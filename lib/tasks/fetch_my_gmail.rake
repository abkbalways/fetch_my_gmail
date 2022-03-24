namespace :gmail_messages do
  require 'gmail'
  require 'rubyXL'
  desc "Fetch Gmail messages periodically"

  task fetch: :environment do
    begin
      workbook = RubyXL::Parser.parse("/home/beryl/Desktop/my_email_data.xlsx")
      email_uid = workbook[0].collect {|row| row[2].value}
      worksheet = workbook[0]
    rescue
      workbook = RubyXL::Workbook.new
      email_uid = []
      worksheet = workbook[0]
      worksheet.add_row(3)
      workbook[0].add_cell(0,1,"S.No.")
      workbook[0].add_cell(0,2,"Email UID")
      workbook[0].add_cell(0,3,"Email From")
      workbook[0].add_cell(0,4,"Email Subject")
      workbook[0].add_cell(0,5,"Email Time")
      workbook[0].add_cell(0,6,1)
    end

    gmail = Gmail.new("noreplyforme2@gmail.com","9997817202")
    i = worksheet.sheet_data[0][6].value.to_i
    count = gmail.inbox.count
    emails = gmail.inbox.emails

    emails.each do |email|
      if !email_uid.include?email.message.message_id
        worksheet.add_row(i)
        worksheet.add_cell(i,1,i.to_s)
        worksheet.add_cell(i,2,email.message.message_id)
        worksheet.add_cell(i,3,email.message.from[0])
        worksheet.add_cell(i,4,email.message.subject)
        worksheet.add_cell(i,5,email.message.date.to_s)
        i = i + 1
        worksheet.add_cell(0,6,i)
      else 
        next
      end
    end
    workbook.write("/home/beryl/Desktop/my_email_data.xlsx")
  end
end