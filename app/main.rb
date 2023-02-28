require "pry"
require "pandoc-ruby"
require "json"
require "spreadsheet"

md_file_path = "tmp/imput.md"
md_text      = File.read(md_file_path)

converter  = PandocRuby.new(md_text, from: :markdown, to: :json)
json_text  = converter.convert
table_data = JSON.parse(json_text)

header_data  = table_data["blocks"][0]["c"][3]
columns_data = table_data["blocks"][0]["c"][4]

header  = header_data.map do |v|
  v[0]["c"][0]["c"]
end
columns = columns_data.map do |v|
  v.map do |_|
    _[0]["c"][0]["c"]
  end
end
base_array = columns.unshift(header)

Spreadsheet.client_encoding = "UTF-8"
book = Spreadsheet::Workbook.new                
sheet = book.create_worksheet(name: "worksheet1")

base_array.each_with_index do |elements, i|
  elements.each_with_index do |e, _i|
    sheet[i, _i] = e
  end
end

book.write("tmp/output.xls")
