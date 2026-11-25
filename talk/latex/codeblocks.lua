-- Maps Markdown fence languages to listings languages. Add your own here.
-- Anything not listed is set as plain (still styled) code.
local map = {
  python = "Python", py = "Python",
  bash = "bash", sh = "bash", shell = "bash", zsh = "bash",
  json = "json", yaml = "yaml", yml = "yaml",
  sql = "SQL", r = "R", c = "C", cpp = "C++", ["c++"] = "C++", java = "Java",
  tex = "TeX", latex = "TeX", html = "HTML", xml = "XML", matlab = "Matlab",
}

function CodeBlock(el)
  local lang = el.classes[1]
  local opt = ""
  if lang and map[lang:lower()] then
    opt = "[language=" .. map[lang:lower()] .. "]"
  end
  return pandoc.RawBlock("latex", "\\begin{lstlisting}" .. opt .. "\n" .. el.text .. "\n\\end{lstlisting}")
end
