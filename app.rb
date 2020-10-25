require 'rubyXL'
require 'prawn'

workbook = RubyXL::Parser.parse("Bewertung 2020HS.xlsx")
worksheet = workbook.worksheets[0]
(3..70).each do |row_num|
  row = worksheet[row_num]

  if row.nil? || row[0].nil? || row[0] == ''
    next
  end

  klass = row[0].value
  group = row[1].value
  pw = row[2].value

  name_indices = [3,4,5]
  authors = name_indices.map { |i| row[i] && row[i].value }.compact

  topics = [
    { start_index: 6, count: 6 },
    { start_index: 20, count: 4 },
    { start_index: 30, count: 4 },
    { start_index: 40, count: 5 }
  ]

  puts authors

  pdf = Prawn::Document.new
  pdf.move_down 20
  pdf.font_size 24
  pdf.text "Zwischenbewertung"
  pdf.move_down 20
  pdf.font_size 10
  pdf.text authors.join(", ")
  pdf.move_down 20
  pdf.stroke_horizontal_rule

  topics.each_with_index do |topic, idx|
    start_index = topic[:start_index]
    next if row[start_index].nil?
    author_lastname = row[start_index].value
    next if author_lastname.nil?
    author = authors.find { |a| a.index(author_lastname) }
    week = idx + 1
    title = worksheet[0][start_index].value
    puts "WOCHE #{week}, Thema: #{title}"
    pdf.stroke_horizontal_rule
    pdf.move_down 10
    pdf.text title
    pdf.text author, align: :right
    pdf.move_down 10
    pdf.stroke_horizontal_rule
    pdf.move_down 20
    (0..(topic[:count]-1)).each do |question_number|
      text_col = start_index + 2 * question_number + 1
      points_col = start_index + 2 * question_number + 2
      question_text = worksheet[1][text_col].value
      question_points = worksheet[1][points_col].value
      grade_text = (row[text_col] && row[text_col].value) || "FEHLER"
      grade_points = (row[points_col] && row[points_col].value) || "FEHLER"
      pdf.text "#{question_text} (#{question_points})", style: :bold
      pdf.text grade_text
      pdf.text grade_points.to_s, align: :right
      pdf.move_down 20
    end
    pdf.text "Total (von 10 Punkten)", style: :bold
    pdf.text (row[start_index + 2 * topic[:count] + 1].value).to_s, align: :right
    pdf.move_down 50
  end

  pdf.encrypt_document(user_password: pw, owner_password: pw)
  pdf.render_file "output/#{klass}-#{group}.pdf"
end
