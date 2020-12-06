require 'rubyXL'
require 'prawn'

workbook = RubyXL::Parser.parse("Bewertung 2020HS.xlsx")
worksheet = workbook.worksheets[0]
(2..70).each do |row_num|
  row = worksheet[row_num]

  if row.nil? || row[0].nil? || row[0] == ''
    next
  end

  klass = row[0].value
  group = row[1].value
  pw = row[2].value

  name_indices = [3,4,5]
  authors = name_indices.map { |i| row[i] && row[i].value }.compact

  # start + 3 + 2*count
  # count = (next - start - 3) / 2
  topics = [
    { start_index: 6, count: 6 },
    { start_index: 21, count: 4 },
    { start_index: 32, count: 4 },
    { start_index: 43, count: 5 },
    { start_index: 56, count: 5 },
    { start_index: 69, count: 5 },
    { start_index: 82, count: 2 }
  ]

  puts authors.inspect

  author_points = {}
  authors.each do |author|
    author_points[author] = { total: 0, max: 0 }
  end

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
      grade_text = (row[text_col] && row[text_col].value) || (raise StandardError, "FEHLER TEXT #{text_col} #{RubyXL::Reference.ind2ref(row_num, text_col)}")
      grade_points = (row[points_col] && row[points_col].value) || (raise StandardError, "FEHLER POINTS #{points_col} #{RubyXL::Reference.ind2ref(row_num, points_col)}")
      pdf.text "#{question_text} (#{question_points})", style: :bold
      pdf.text grade_text
      pdf.text grade_points.to_s, align: :right
      pdf.move_down 20
    end

    max = row[start_index + 2 * topic[:count] + 1].value
    total = row[start_index + 2 * topic[:count] + 2].value

    pdf.text "Total (von #{max} Punkten)", style: :bold
    pdf.text total.to_s, align: :right
    if idx < topic.count
      pdf.start_new_page
      pdf.move_down 80
    end

    if author_points[author]
      author_points[author][:max] += max
      author_points[author][:total] += total
    end
  end

  pdf.move_down 100

  authors.each do |author|
    pdf.stroke_horizontal_rule
    pdf.move_down 20
    sum_max = author_points[author][:max]
    sum_total = author_points[author][:total]
    pdf.text author
    pdf.text "Punkte: #{sum_total} von #{sum_max}"
    if sum_max > 0
      pdf.text "Note: #{(sum_total / sum_max.to_f * 5 + 1).round(1)}"
    else
      pdf.text "Note: 0"
    end
    pdf.move_down 20
  end
  pdf.stroke_horizontal_rule

  pdf.encrypt_document(user_password: pw, owner_password: pw)
  pdf.render_file "output/#{klass}-#{group}.pdf"
end
