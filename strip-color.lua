-- strip-color.lua — Pandoc filter: remove ALL color from output
-- This ensures the generated .docx has only black text (no colored spans/links/headings)

local function strip_color(el)
  -- Strip color from Span
  if el.t == 'Span' then
    el.attributes['color'] = nil
    el.attributes['background-color'] = nil
    return el
  end
  -- Strip color from Link
  if el.t == 'Link' then
    el.attributes['color'] = nil
    return el
  end
  return nil
end

-- Handle Code blocks (sometimes colored)
local function strip_code(el)
  if el.t == 'CodeBlock' or el.t == 'Code' then
    el.attributes['color'] = nil
    return el
  end
  return nil
end

return {
  { Span = strip_color },
  { Link = strip_color },
  { Code = strip_code },
  { CodeBlock = strip_code },
}
