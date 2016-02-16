require 'roo'
require 'json'
# Require the WIN32OLE library
require 'win32ole'

class EXCEL_CONST
end
excel = WIN32OLE.new('Excel.Application')
WIN32OLE.const_load(excel, EXCEL_CONST)




class Meetings
  attr_accessor :meetings_hash
  
  def read_excel(file = 'C:\test\Template - Database of stakeholder input.xlsx' )
    xlsx = Roo::Excelx.new(file)
    sheet = xlsx.sheet('Data')
    #Set Headers in Sheet and create hash.  
    headers =  sheet.row(1)
    @meetings_hash = sheet.parse(header_search: headers)
    @meetings_hash.each_with_index do |row,i|
      attendees = row["Attendees"]
      array = attendees.split(/\r?\n/)
      people = []
      array.each do |person|
        array = person.split(':') 
        person = Hash.new()
        person["short name"] =  array[0].strip unless array[0] == nil
        person["full name"] =   array[1].strip unless array[1] == nil
        person["organization"]= array[2].strip unless array[2] == nil
        people << person
      end
      row["Attendees"] = people
      #Global Hashtags
      global_tags = []
      row["Global Tags"].split(/\r?\n/).each { |tag| global_tags << tag.strip } unless row["Global Tags"] == nil or row["Global Tags"].strip == ""
      
      tweets = []
      tweet = Hash.new()
      st = row["Significant Tweets"]
      array = st.split(/(?<=[\r?\n])/) unless st == nil
      array.each do |quote|
        quote.strip!
        unless quote == "" or quote == nil
          #break quote into tags and quote parts. $1,$2. 
          quote =~ /(^@.*?):(.*)$/
          first =  $1
          raw_quote =  $2
          if raw_quote.nil? or raw_quote.strip == ""
            next
          else
            raw_quote.strip!
          end
          #get people short names from $1
          people = []
          people = first.scan(/\B@\w+/) unless first == nil
          #If no people are found, or @ALL is indicated add all attendees to quote
          if people.nil? or people[0].nil?  or people[0].downcase == "@ALL".downcase
            row["Attendees"].each {|person| people << person['short name']}
          end
          #Get Hashtags
          tags = first.scan(/\B#\w+/) + global_tags unless first.nil?
          tags = tags.uniq
          tweet["people"] = people
          tweet["hashtags"] =  tags
          tweet["quote"] =  raw_quote
          tweets << tweet
        end
      end
      row["Significant Tweets"] = tweets
    end
    @meetings_hash.shift
    puts JSON.pretty_generate(@meetings_hash)
    return @meetings_hash
  end

  def write_excel(file = 'C:\test\new.xlsx')
    self.read_excel()
    
    #Create Header based on first JSON
    header_array = []
    @meetings_hash.each do |document|
      document.each do |k,v|
        header_array << k
        header_array = header_array.uniq
      end
    end

    
    # Create an instance of the Excel application object
    xl = WIN32OLE.new('Excel.Application')
    # Make Excel visible
    xl.Visible = 1
    # Add a new Workbook object
    wb = xl.Workbooks.Add
    # Get the first Worksheet
    ws = wb.Worksheets(1)
    # Set the name of the worksheet tab
    ws.Name = 'Data'
    
    #Add header
    header_array.each_with_index do |header, c|
      ws.Cells(1, c + 1).Value = header
    end
    # For each row in the data set
    @meetings_hash.each_with_index do |document , r|
      header_array.each_with_index do |header, c|
        case header
        when "Attendees"
          document[header].each_with_index do |person,i|
            personstring = "#{person['short name']} : #{person['full name']} : #{person['organization']}"
            ws.Cells(r + 2, c + 1).Value = "#{ws.Cells(r + 2, c + 1).Value} #{personstring}\n"
          end

        when "Significant Tweets"
          
          document[header].each_with_index do |tweet,i|
            tags =""
            tweet['people'].each {|person| tags = "#{tags} #{person}" } 
            tweet["hashtags"].each {|tag| tags = "#{tags} #{tag}" } 
            quote =tweet["quote"] unless tweet.nil?
            if tweet == document[header].last
              ws.Cells(r + 2, c + 1).Value = "#{ws.Cells(r + 2, c + 1).Value}#{tags} : #{quote}\n\n"
            else
              ws.Cells(r + 2, c + 1).Value = "#{ws.Cells(r + 2, c + 1).Value}#{tags} : #{quote}"
            end
            #ws.Cells(r + 2, c + 1).Characters(1,tags.size).Font.Bold = true
            #Bold tags
          end
        else
          ws.Cells(r + 2, c + 1).Value = document[header].to_s
        end
      end

      
    end
    
    #apply style
    tbl = ws.ListObjects.Add(nil, ws.Range("A1","N10"), nil , 1)
    tbl.TableStyle = "TableStyleMedium11"
    ws.Rows.VerticalAlignment = EXCEL_CONST::XlTop
    ws.Columns.HorizontalAlignment = EXCEL_CONST::XlLeft
    ws.Columns.WrapText = true
    ws.Rows(1).VerticalAlignment = EXCEL_CONST::XlCenter
    ws.Rows(1).HorizontalAlignment = EXCEL_CONST::XlCenter
    
    [1,2,3,4].each do |col|
      ws.Columns(col).NumberFormat = "yyyy/mm/dd"
      ws.Columns(col).HorizontalAlignment = -4108
      ws.Columns(col).VerticalAlignment = -4108
      ws.Columns(col).Orientation = 90
      ws.Columns(col).ColumnWidth = 3.0
    end
    #wb.SaveAs('c:\temp\workbook.xls')
    # Close the workbook
    wb.Close
    # Quit Excel
    xl.Quit
  end
 

  
end





Meetings.new.write_excel