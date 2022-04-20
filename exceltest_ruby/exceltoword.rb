require "creek"
require "openxml/docx"

def word(item, file)
  package = OpenXml::Docx::Package.new
  include OpenXml::Docx::Elements
  line = item.split("\n")
  line.each do |lorem|
    text = Text.new(lorem)
    run = Run.new
    run << text
    paragraph = Paragraph.new
    paragraph << run
    package.document << paragraph
  end
  section_properties = SectionProperties.new
  font_size = OpenXml::Docx::Properties::FontSize.new(19)
  font_name = OpenXml::Docx::Properties::Font.new("宋体")
  section_properties << font_size
  section_properties << font_name
  package.document << section_properties
  package.save(file)
end

def joinlist(item, len)
  letters = ["B", "C", "D", "E", "F", "G", "H"]
  item[0] = item[0].to_s.delete("\n") + "\n"
  if item[1] && item[1].to_s.strip != ""
    item[1] = "A." + item[1].to_s.delete("\n")
  else
    item[1] = ""
  end
  if len == 3
    item[len - 1] = " 答案：" + item[len - 1].to_s.delete("\n") + "\n"
    return item.join
  end
  if len > 3
    for i in 0...(len - 3)
      if item[i + 2] && item[i + 2].to_s.strip != ""
        item[i + 2] = " " + letters[i] + "." + item[i + 2].to_s.delete("\n")
      else
        item[i + 2] = ""
      end
    end
    item[len - 1] = " 答案：" + item[len - 1].to_s.delete("\n") + "\n"
    return item.join
  end
end

def process(items)
  questionindex = []
  question = []
  answer = []
  choice = []
  for item in items
    if /题目/ =~ item
      question.push(items.index(item))
    end
    if /[ABCDEFGH]/i =~ item
      choice.push(items.index(item))
    end
    if /答案/ =~ item
      answer.push(items.index(item))
    end
  end
  if question.any?
    questionindex.push(question[0])
  end
  if choice.any?
    questionindex.concat(choice)
  end
  if answer.any?
    questionindex.push(answer[0])
  end
  if questionindex.any?
    return questionindex, questionindex.length
  end
  return nil, nil
end

def getcontentofsheet(sheet)
  temp = []
  sheet.rows.each_slice(1) do |row|
    temp.push(row[0].values)
  end
  if temp.any?
    return temp.length, temp[0], temp[0].length, temp
  end
  return nil, nil, nil, nil
end

data = Creek::Book.new "safety.xlsx"
sheetnumber = data.sheets.length
total = Array.new(sheetnumber)
table = Array.new(sheetnumber)
nrows = Array.new(sheetnumber)
ncols = Array.new(sheetnumber)
sum = Array.new(sheetnumber)

for i in 0...sheetnumber
  table[i] = data.sheets[i]
  nrows[i], firstrow, ncols[i], content = getcontentofsheet(table[i])
  if content
    sum[i] = []
    temp, fuck = process(firstrow)
    if temp
      for j in 0...(nrows[i] - 1)
        everyrow = []
        index = "(" + (j + 1).to_s + ")"
        for k in temp
          everyrow.push(content[j + 1][k.to_i])
        end
        everyrowlist = index + joinlist(everyrow, fuck)
        sum[i].push(everyrowlist)
      end
      total[i] = sum[i].join
      title = table[i].name.to_s + "Ruby.docx"
      word(total[i], title)
    end
  end
end
