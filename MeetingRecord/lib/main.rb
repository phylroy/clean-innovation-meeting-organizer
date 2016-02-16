require 'json'
require 'sequel'

require 'sequel'



class MeetingRecord

  
  attr_accessor :date, 
    :title , 
    :location, 
    :meeting_type, 
    :people,
    :context,
    :objectives,
    :backround,
    :considerations,
    :tweets,
    :actions,
    :full_notes
  def initialize()
    #make some variables hashes. 
    @tweets = {}
    @actions = {}
    @people = {}
  end
  
  def import_from_excel_row(array)
    #Open Excel file. 
    #Read @title , 
    location, 
    :meeting_type, 
    :people,
    :context,
    :objectives,
    :backround,
    :considerations,
    :tweets,
    :actions,
    :full_notes
  end
  
  def read_from_word_template(file)
    
  end
  
end

class MeetingRecordList
  attr_accessor :meetingList
  def read_from_excel(file)
    #Open Excel file. 
    #Read @title , 
    location, 
    :meeting_type, 
    :people,
    :context,
    :objectives,
    :backround,
    :considerations,
    :tweets,
    :actions,
    :full_notes
  end
  
  def initialize()
    @meetingList = []
  end
end




require 'docx'

class MeetingRecord
  
  attr_accessor :data
  SECTIONS = 
    [
    "MEETING TITLE",
    "MEETING TYPE",
    "DATE",
    "ATTENDEES",
    "CONTEXT DESCRIPTION",
    "CLEAN INNOVATION  AVAILABLE THEMES HASHTAGS",
    "SIGNIFICANT POINTS RAISED",
    "INTERNAL ACTION ITEMS",
    "KEYWORDS",
    "FULL TEXT DUMP"
  ]



  def get_content(doc,index)
    lines = []
    (index+1..doc.size+1).each do |line_no|
      if SECTIONS.include?( doc[line_no].to_s.strip.upcase)
        return lines
      else
        if doc[line_no].to_s.strip != ""
          lines << doc[line_no].to_s.strip
        end
      end
    end
    return lines
  end

  def initialize(file)
    @data = []
    @file = file
    # Create a Docx::Document object for our existing docx file
    doc = Docx::Document.open(file).to_s.split(/\n+/)
    # Retrieve and display paragraphs
      doc.each_with_index do |line, index|
        if SECTIONS.include? (line.upcase.strip)
          content = get_content( doc, index )
          instance_variable_set("@#{line.downcase.tr(" ", "_")}", content )
          @data << content
          puts "@#{line.downcase.tr(" ", "_")} = #{content}"
        end
      end
  end
end

mr = MeetingRecord.new('C:/temp/Meeting Summary Template.docx')

# Require the WIN32OLE library
require 'win32ole'



# Create an instance of the Excel application object
xl = WIN32OLE.new('Excel.Application')
# Make Excel visible
xl.Visible = 1
# Add a new Workbook object
wb = xl.Workbooks.Add
# Get the first Worksheet
ws = wb.Worksheets(1)
# Set the name of the worksheet tab
ws.Name = 'Stakeholder Feedback'
# For each row in the data set
mr.data.each_with_index do |row, r|
  # For each field in the row
  row.each_with_index do |field, c|
      # Write the data to the Worksheet
      ws.Cells(r+1, c+1).Value = field.to_s
  end
end
# Save the workbook

tbl = ws.ListObjects.Add(nil, ws.Range("A1","J5"), nil , 1)
tbl.TableStyle = "TableStyleMedium11"


wb.SaveAs('c:\temp\workbook.xls')
# Close the workbook
wb.Close
# Quit Excel
xl.Quit
