local M = {}

local PYTHON = "C:\\software\\anaconda3\\python.exe"
local WORD_EXTS = { docx = true, doc = true, docm = true, dotx = true, dotm = true, dot = true, rtf = true }
local PPT_EXTS = { pptx = true, pptm = true, ppt = true, ppsx = true, ppsm = true, pps = true,
	potx = true, potm = true, pot = true }

local function plugin_file(name)
	local config = os.getenv("YAZI_CONFIG_HOME") or (os.getenv("APPDATA") .. "/yazi/config")
	return config .. "/plugins/omni-previewer.yazi/" .. name
end

local function ext_of(url)
	return tostring(url):lower():match("%.(%a+)$") or ""
end

-- ==================== Word page-image pipeline ====================

local state_get = ya.sync(function(state, key)
	local info = state.entries and state.entries[key]
	local pending = state.pending and state.pending[key]
	return state.text_mode, state.failed and state.failed[key], info, pending
end)

local state_ready = ya.sync(function(state, key, info)
	state.entries = state.entries or {}
	state.entries[key] = info
	state.pending = state.pending or {}
	state.pending[key] = nil
	state.thumbs = state.thumbs or {}
	state.thumbs[key] = nil
end)

local state_pending = ya.sync(function(state, key)
	state.pending = state.pending or {}
	state.pending[key] = os.time()
end)

local state_fail = ya.sync(function(state, key)
	state.failed = state.failed or {}
	local prev = state.failed[key]
	state.failed[key] = { t = os.time(), n = (type(prev) == "table" and prev.n or 0) + 1 }
	state.pending = state.pending or {}
	state.pending[key] = nil
end)

local state_thumb_get = ya.sync(function(state, key)
	return state.thumbs and state.thumbs[key]
end)

local state_thumb_set = ya.sync(function(state, key, path)
	state.thumbs = state.thumbs or {}
	state.thumbs[key] = path
end)

local state_retry_arm = ya.sync(function(state, key)
	local f = state.failed and state.failed[key]
	if type(f) == "table" then
		f.t = os.time()
	end
end)

local state_unfail = ya.sync(function(state)
	state.failed = {}
end)

local toggle = ya.sync(function(state)
	local hovered = cx.active.current.hovered
	local ext = hovered and ext_of(hovered.url) or ""
	if not (hovered and (WORD_EXTS[ext] or PPT_EXTS[ext] or ext == "md" or ext == "markdown")) then
		return nil
	end
	state.text_mode = not state.text_mode
	state.failed = {}
	state.pending = {}
	local skip = 0
	pcall(function()
		skip = cx.active.preview.skip or 0
	end)
	return state.text_mode, skip
end)

local function identity(job)
	return tostring(job.file.url) .. ":" .. tostring(job.file.cha.mtime) .. ":" .. tostring(job.file.cha.len)
end

-- Estimate the pane's pixel long-edge (cell ≈ 10x20 px) with 2x oversampling.
-- Far smaller than the old max_width*16 default -> much faster JPEG decode.
local function pane_edge(job)
	return math.max(800, math.min(3200, math.max(job.area.w, job.area.h * 2) * 20))
end

local edge_seen = ya.sync(function(state, e)
	if e then
		state.last_edge = e
	end
	return state.last_edge
end)

local function show_image(job, image, page)
	if page ~= job.skip then
		return ya.emit("peek", { page, only_if = job.file.url, upper_bound = true })
	end
	local _, err = ya.image_show(Url(image), job.area)
	ya.preview_widget(job, err)
end

local function ansi_lines(job, text, nowrap)
	return ui.lines(text, { ansi = true, tab_size = rt.preview.tab_size, width = job.area.w,
		wrap = nowrap and ui.Wrap.NO or ui.Wrap.YES })
end

local function ansi_page(job, lines)
	local limit = job.area.h
	if job.skip > 0 and job.skip >= #lines then
		return nil, math.max(0, #lines - limit)
	end
	local page = {}
	for i = job.skip + 1, math.min(#lines, job.skip + limit) do
		page[#page + 1] = lines[i]
	end
	return page
end

local function ansi_peek(job, text)
	local page, reskip = ansi_page(job, ansi_lines(job, text))
	if page == nil then
		return ya.emit("peek", { reskip, only_if = job.file.url, upper_bound = true })
	end
	ya.preview_widget(job, { ui.Clear(job.area), ui.Text(page):area(job.area) })
end

local function text_cache(job)
	local cha = job.file.cha
	local h = 5381
	local url = tostring(job.file.url)
	for i = 1, #url do
		h = (h * 33 + url:byte(i)) % 4294967296
	end
	return string.format("%s\\yazi\\preview-cache\\%08x-%x-%x-%dx%d-v9.ansi",
		os.getenv("LOCALAPPDATA") or "", h, cha.len or 0, math.floor(cha.mtime or 0),
		job.area.w, job.area.h)
end

local function read_file(path)
	local ok, data = pcall(function()
		local fh = io.open(path, "rb")
		if not fh then
			return nil
		end
		local d = fh:read("*a")
		fh:close()
		return d
	end)
	if ok then
		return data
	end
end

local function script_peek(job, script)
	local cache = text_cache(job)
	local text = read_file(cache)
	if text == nil then
		local output = Command(PYTHON)
			:arg({ "-X", "utf8", plugin_file(script), tostring(job.file.path), cache, tostring(job.area.w), tostring(job.area.h) })
			:output()
		if not output or not output.status.success then
			return nil
		end
		text = output.stdout
	end
	return text
end

local function word_fallback(job)
	local ext = ext_of(job.file.url)
	if ext == "docx" then
		local text = script_peek(job, "docx_text.py")
		if text then
			return ansi_peek(job, text)
		end
		return require("docx-preview"):peek(job)
	end
	if WORD_EXTS[ext] then
		-- Legacy OLE Word formats: text comes from the persistent server
		-- (Word Content.Text), no OOXML fast path exists for them. Text
		-- requests jump the server's request queue, and the result lands
		-- in the ansi cache - the COM open is paid once per file.
		local text = script_peek(job, "doc_text.py")
		if text then
			return ansi_peek(job, text)
		end
	end
	if PPT_EXTS[ext] then
		local text = script_peek(job, "pptx_text.py")
		if text then
			return ansi_peek(job, text)
		end
	end
	ya.preview_widget(job, { ui.Clear(job.area), ui.Text({
		ui.Line("Page image will appear when ready (text preview is unavailable for this format)"),
	}):area(job.area) })
end

local function word_probe(job, key, edge)
	local output = Command(PYTHON)
		:arg({ "-X", "utf8", plugin_file("render.py"), tostring(job.file.path), tostring(job.skip), tostring(edge), "--probe" })
		:output()
	local result = output and output.status.success and ya.json_decode(output.stdout) or nil
	if result and result.image then
		state_ready(key, { dir = result.dir, edge = result.edge, pages = result.pages, png = result.png })
		return show_image(job, result.image, result.page)
	end
	if result and result.thumb then
		-- Conversion is still in-flight; pin the thumbnail so later peeks
		-- (while pending) re-show it without spawning a probe each time.
		state_thumb_set(key, result.thumb)
		return show_image(job, result.thumb, job.skip)
	end
	if output and output.stderr and output.stderr:match("PREVFAILED") then
		state_fail(key)
	end
	return word_fallback(job)
end

local function spawn_worker(job)
	-- Detached --notify render: converts, then emits plugin "refresh" which
	-- clears failed state and re-peeks. Best effort; spawn may be unavailable.
	pcall(function()
		Command(PYTHON)
			:arg({ "-X", "utf8", plugin_file("render.py"), tostring(job.file.path), "0",
				tostring(pane_edge(job)), "--notify", tostring(job.file.url) })
			:spawn()
	end)
end

local function word_peek(job)
	local key = identity(job)
	local text_mode, failed, info, pending = state_get(key)
	if text_mode then
		return word_fallback(job)
	end
	if failed then
		-- Transient conversion failures used to stick for the whole session:
		-- retry up to 3 times with a 20s cooldown; a success notifies "refresh"
		-- which clears the failure and re-peeks automatically.
		if failed.n <= 3 and os.time() - failed.t >= 20 then
			state_retry_arm(key)
			spawn_worker(job)
		end
		return word_fallback(job)
	end

	local edge = pane_edge(job)
	edge_seen(edge)
	if info then
		local page = math.min(job.skip, info.pages - 1)
		local image = info.png and (info.dir .. "/page-" .. page .. ".png")
			or (info.dir .. "/page-" .. page .. "-" .. info.edge .. ".jpg")
		if fs.cha(Url(image)) then
			return show_image(job, image, page)
		end
		return word_probe(job, key, edge)
	end
	if pending then
		-- pending is only a timestamped guess that a worker is in-flight; if it
		-- is stale the worker died without notifying - kick a fresh one.
		if os.time() - pending > 60 then
			state_pending(key)
			spawn_worker(job)
		end
		local thumb = state_thumb_get(key)
		if thumb and fs.cha(Url(thumb)) then
			return show_image(job, thumb, job.skip)
		end
		return word_fallback(job)
	end

	local output = Command(PYTHON)
		:arg({ "-X", "utf8", plugin_file("render.py"), tostring(job.file.path), tostring(job.skip), tostring(edge), "--probe" })
		:output()
	local result = output and output.status.success and ya.json_decode(output.stdout) or nil
	if result and result.image then
		state_ready(key, { dir = result.dir, edge = result.edge, pages = result.pages, png = result.png })
		return show_image(job, result.image, result.page)
	end
	if result and result.thumb then
		state_thumb_set(key, result.thumb)
		state_pending(key)
		return show_image(job, result.thumb, job.skip)
	end

	state_pending(key)
	word_fallback(job)
end

local function word_seek(job)
	local text_mode, failed, info, pending = state_get(identity(job))
	local hovered = cx.active.current.hovered
	if not (hovered and hovered.url == job.file.url) then
		return
	end
	-- Line-unit scrolling when showing text (text mode, failure, or the
	-- fallback shown while conversion is pending); page-unit otherwise.
	if text_mode or failed or (not info and pending) then
		ya.emit("peek", { math.max(0, cx.active.preview.skip + job.units), only_if = job.file.url })
	else
		ya.emit("peek", { math.max(0, cx.active.preview.skip + ya.clamp(-1, job.units, 1)), only_if = job.file.url })
	end
end

-- ==================== Grid table pipeline ====================

local GRID_SCRIPTS = { xlsx = "xlsx.py", xls = "xlsx.py", csv = "table.py" }

local function grid_peek(job)
	local text = script_peek(job, GRID_SCRIPTS[ext_of(job.file.url)])
	if text == nil then
		return require("empty").msg(job, "Table render failed")
	end
	ansi_peek(job, text)
end

-- ==================== Markdown pipeline ====================

local MD_EXTS = { md = true, markdown = true }

local function md_peek(job)
	local text_mode = state_get(identity(job))
	if text_mode then
		-- Drop any overlay image still painted from rendered mode (best effort;
		-- the builtin code peeker cannot erase it for us).
		pcall(function()
			ya.image_hide()
		end)
		return require("code"):peek(job)
	end
	local raw = script_peek(job, "md.py")
	if raw == nil then
		return require("code"):peek(job)
	end
	local manifest = ya.json_decode(raw)
	if type(manifest) ~= "table" or not manifest.text then
		return ansi_peek(job, raw)
	end
	-- Text is pre-wrapped to pane width by md.py, so manifest line numbers map
	-- 1:1 to screen rows: overlay each media image into its placeholder sub-rect.
	-- ui.Clear wipes leftovers of earlier overlay cells that scrolled out.
	local lines = ansi_lines(job, manifest.text, true)
	local page, reskip = ansi_page(job, lines)
	if page == nil then
		return ya.emit("peek", { reskip, only_if = job.file.url, upper_bound = true })
	end
	ya.preview_widget(job, { ui.Clear(job.area), ui.Text(page):area(job.area) })
	local limit = job.area.h
	for _, m in ipairs(manifest.media or {}) do
		local top, h = tonumber(m.line) or -1, tonumber(m.lines) or 0
		if m.path and top >= job.skip and top + h <= job.skip + limit and h >= 2 then
			ya.image_show(Url(m.path), ui.Rect {
				x = job.area.x, y = job.area.y + top - job.skip,
				w = job.area.w, h = h,
			})
		end
	end
end

-- ==================== Dispatch ====================

local PAGE_EXTS = { pdf = true }
for ext in pairs(WORD_EXTS) do
	PAGE_EXTS[ext] = true
end
for ext in pairs(PPT_EXTS) do
	PAGE_EXTS[ext] = true
end

function M:peek(job)
	local ext = ext_of(job.file.url)
	if PAGE_EXTS[ext] then
		return word_peek(job)
	elseif GRID_SCRIPTS[ext] then
		return grid_peek(job)
	elseif MD_EXTS[ext] then
		return md_peek(job)
	end
end

function M:preload(job)
	local file = job.file
	if not file or not PAGE_EXTS[ext_of(file.url)] then
		return true
	end
	local edge = edge_seen() or 1600
	Command(PYTHON)
		:arg({ "-X", "utf8", plugin_file("render.py"), tostring(file.path or file.url), "0",
			tostring(edge), "--notify", tostring(file.url) })
		:output()
	return true
end

function M:seek(job)
	if PAGE_EXTS[ext_of(job.file.url)] then
		return word_seek(job)
	end
	return require("code"):seek(job)
end

function M:entry(job)
	local raw = job and job.args
	if type(raw) == "table" then
		raw = table.concat(raw, " ")
	end
	raw = tostring(raw or "")
	local parts = {}
	local pat = raw:find("|", 1, true) and "[^|]+" or "%S+"
	for p in raw:gmatch(pat) do
		parts[#parts + 1] = p
	end
	if parts[1] == "refresh" then
		state_unfail()
		ya.emit("peek", { tonumber(parts[#parts]) or 0, force = true })
		return
	end

	local text_mode, skip = toggle()
	if text_mode == nil then
		return
	end
	ya.notify({ title = "Preview", content = text_mode and "Text mode" or "Page image mode", timeout = 2 })
	ya.emit("peek", { tonumber(skip) or 0, force = true })
end

return M
