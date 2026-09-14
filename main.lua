local M = {}

local PYTHON = "C:\\software\\anaconda3\\python.exe"
local WORD_EXTS = { docx = true, doc = true, docm = true, dotx = true, dotm = true, rtf = true }

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
end)

local state_pending = ya.sync(function(state, key)
	state.pending = state.pending or {}
	state.pending[key] = true
end)

local state_fail = ya.sync(function(state, key)
	state.failed = state.failed or {}
	state.failed[key] = true
	state.pending = state.pending or {}
	state.pending[key] = nil
end)

local state_unfail = ya.sync(function(state)
	state.failed = {}
end)

local toggle = ya.sync(function(state)
	local hovered = cx.active.current.hovered
	if not (hovered and WORD_EXTS[ext_of(hovered.url)]) then
		return nil
	end
	state.text_mode = not state.text_mode
	state.failed = {}
	state.pending = {}
	return state.text_mode
end)

local function identity(job)
	return tostring(job.file.url) .. ":" .. tostring(job.file.cha.mtime) .. ":" .. tostring(job.file.cha.len)
end

local function show_image(job, image, page)
	if page ~= job.skip then
		return ya.emit("peek", { page, only_if = job.file.url, upper_bound = true })
	end
	local _, err = ya.image_show(Url(image), job.area)
	ya.preview_widget(job, err)
end

local function ansi_peek(job, text)
	local lines = ui.lines(text, { ansi = true, tab_size = rt.preview.tab_size, width = job.area.w, wrap = ui.Wrap.YES })
	local limit = job.area.h
	if job.skip > 0 and job.skip >= #lines then
		return ya.emit("peek", { math.max(0, #lines - limit), only_if = job.file.url, upper_bound = true })
	end
	local page = {}
	for i = job.skip + 1, math.min(#lines, job.skip + limit) do
		page[#page + 1] = lines[i]
	end
	ya.preview_widget(job, ui.Text(page):area(job.area))
end

local function text_cache(job)
	local cha = job.file.cha
	local h = 5381
	local url = tostring(job.file.url)
	for i = 1, #url do
		h = (h * 33 + url:byte(i)) % 4294967296
	end
	return string.format("%s\\yazi\\preview-cache\\%08x-%x-%x.ansi",
		os.getenv("LOCALAPPDATA") or "", h, cha.len or 0, math.floor(cha.mtime or 0))
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
			:arg({ "-X", "utf8", plugin_file(script), tostring(job.file.path), cache })
			:output()
		if not output or not output.status.success then
			return nil
		end
		text = output.stdout
	end
	return text
end

local function word_fallback(job)
	if ext_of(job.file.url) == "docx" then
		local text = script_peek(job, "docx_text.py")
		if text then
			return ansi_peek(job, text)
		end
		return require("docx-preview"):peek(job)
	end
	ya.preview_widget(job, ui.Text({ ui.Line("Page image will appear when ready (text preview is unavailable for this format)") }):area(job.area))
end

local function word_probe(job, key, edge)
	local output = Command(PYTHON)
		:arg({ "-X", "utf8", plugin_file("render.py"), tostring(job.file.path), tostring(job.skip), tostring(edge), "--probe" })
		:output()
	local result = output and output.status.success and ya.json_decode(output.stdout) or nil
	if result and result.image then
		state_ready(key, { dir = result.dir, edge = result.edge, pages = result.pages })
		return show_image(job, result.image, result.page)
	end
	if output and output.stderr and output.stderr:match("PREVFAILED") then
		state_fail(key)
	end
	return word_fallback(job)
end

local function word_peek(job)
	local key = identity(job)
	local text_mode, failed, info, pending = state_get(key)
	if text_mode or failed then
		return word_fallback(job)
	end

	local edge = math.max(600, math.min(3200, math.max(rt.preview.max_width, rt.preview.max_height) * 16))
	if info then
		local page = math.min(job.skip, info.pages - 1)
		local image = info.dir .. "/page-" .. page .. "-" .. info.edge .. ".jpg"
		if fs.cha(Url(image)) then
			return show_image(job, image, page)
		end
		return word_probe(job, key, edge)
	end
	if pending then
		return word_probe(job, key, edge)
	end

	local output = Command(PYTHON)
		:arg({ "-X", "utf8", plugin_file("render.py"), tostring(job.file.path), tostring(job.skip), tostring(edge), "--probe" })
		:output()
	local result = output and output.status.success and ya.json_decode(output.stdout) or nil
	if result and result.image then
		state_ready(key, { dir = result.dir, edge = result.edge, pages = result.pages })
		return show_image(job, result.image, result.page)
	end

	state_pending(key)
	word_fallback(job)
end

local function word_seek(job)
	local text_mode, failed = state_get(identity(job))
	if text_mode or failed then
		return require("docx-preview"):seek(job)
	end
	local hovered = cx.active.current.hovered
	if hovered and hovered.url == job.file.url then
		ya.emit("peek", { math.max(0, cx.active.preview.skip + ya.clamp(-1, job.units, 1)), only_if = job.file.url })
	end
end

-- ==================== Grid table pipeline ====================

local GRID_SCRIPTS = { xlsx = "xlsx.py", csv = "table.py" }

local function grid_peek(job)
	local text = script_peek(job, GRID_SCRIPTS[ext_of(job.file.url)])
	if text == nil then
		return require("empty").msg(job, "Table render failed")
	end
	ansi_peek(job, text)
end

-- ==================== Dispatch ====================

local PAGE_EXTS = { pdf = true }
for ext in pairs(WORD_EXTS) do
	PAGE_EXTS[ext] = true
end

function M:peek(job)
	local ext = ext_of(job.file.url)
	if PAGE_EXTS[ext] then
		return word_peek(job)
	elseif GRID_SCRIPTS[ext] then
		return grid_peek(job)
	end
end

function M:preload(job)
	local file = job.file
	if not file or not PAGE_EXTS[ext_of(file.url)] then
		return true
	end
	Command(PYTHON)
		:arg({ "-X", "utf8", plugin_file("render.py"), tostring(file.path or file.url), "0", "2000", "--notify", tostring(file.url) })
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
	local text_mode = toggle()
	if text_mode == nil then
		return
	end
	ya.notify({ title = "Word preview", content = text_mode and "Text mode" or "Page image mode", timeout = 2 })
	ya.emit("peek", { 0, force = true })
end

return M
