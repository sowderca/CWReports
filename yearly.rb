require 'business_time'
require 'connect-stoopid'
require 'axlsx'
require 'mailgun'
require 'date'
require 'time'
require 'csv'
#require 'to_csv'


#Mailgun 
mg_client = Mailgun::Client.new("key-70f741f0b10b167ad20b09f841151af8")
mb_obj = Mailgun::MessageBuilder.new

def remove_file(file)
  File.delete(file)
end

#class Array
#    def to_csv(csv_filename="hash.csv")
#    CSV.open(csv_filename, "wb") do |csv|
#        csv << first.keys
#        self.each do |hash|
#            csv << hash.value
#        end
#        end
#    end
#end
=begin vars
Set the dates for the report
=end
start_date = Date.parse("2014-04-01")
end_date = Date.parse("2015-05-01")
comps = Array.new #array for new dates
#gets only the weekdays from the start date untill the end date.
days = start_date.business_dates_until(end_date)

days.each do | modify | 
  comps << modify.iso8601
end
array = Array.new
comps.each do |day|
	@morning = "[#{day}T07:00:00-04:00]"
	@evening = "[#{day}T18:00:00-04:00]"
	cw_client = ConnectStoopid::ReportingClient.new('connect.vc3.com', 'vc3', 'dattoCW', 'mushySh@pe83')
    
    @fcr_c = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@morning} and date_closed < #{@evening} and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'witherst' or closed_by = 'washingt' or closed_by = 'sleeperk' or closed_by = 'chandles' or closed_by = 'keddingl' or closed_by = 'williamm' or closed_by = 'lettswil' or closed_by = 'greencra' or closed_by = 'smitheli') and Board_Name = 'FCR'")
	@infa_c = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@morning} and date_closed < #{@evening} and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'witherst' or closed_by = 'washingt' or closed_by = 'sleeperk' or closed_by = 'chandles' or closed_by = 'keddingl' or closed_by = 'williamm' or closed_by = 'lettswil' or closed_by = 'greencra' or closed_by = 'smitheli') and Board_Name = 'Infrastructure'")
    @comm_c = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@morning} and date_closed < #{@evening} and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'witherst' or closed_by = 'washingt' or closed_by = 'sleeperk' or closed_by = 'chandles' or closed_by = 'keddingl' or closed_by = 'williamm' or closed_by = 'lettswil' or closed_by = 'greencra' or closed_by = 'smitheli') and (Board_Name = 'SC Comm Internal' or Board_Name  ='SC Comm Services')") 
    @sc_c = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@morning} and date_closed < #{@evening} and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'witherst' or closed_by = 'washingt' or closed_by = 'sleeperk' or closed_by = 'chandles' or closed_by = 'keddingl' or closed_by = 'williamm' or closed_by = 'lettswil' or closed_by = 'greencra' or closed_by = 'smitheli') and (Board_Name = 'SC Govt Internal' or Board_Name = 'SC Govt Internal')")
    @nc_c = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@morning} and date_closed < #{@evening} and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'witherst' or closed_by = 'washingt' or closed_by = 'sleeperk' or closed_by = 'chandles' or closed_by = 'keddingl' or closed_by = 'williamm' or closed_by = 'lettswil' or closed_by = 'greencra' or closed_by = 'smitheli') and (Board_Name = 'NC Govt Internal' or Board_Name = 'NC Govt Services')")
    @ga_c = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@morning} and date_closed < #{@evening} and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'witherst' or closed_by = 'washingt' or closed_by = 'sleeperk' or closed_by = 'chandles' or closed_by = 'keddingl' or closed_by = 'williamm' or closed_by = 'lettswil' or closed_by = 'greencra' or closed_by = 'smitheli') and (Board_Name = 'GA Govt Internal' or Board_Name = 'GA Govt Services')")
    @lt_c = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@morning} and date_closed < #{@evening} and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'witherst' or closed_by = 'washingt' or closed_by = 'sleeperk' or closed_by = 'chandles' or closed_by = 'keddingl' or closed_by = 'williamm' or closed_by = 'lettswil' or closed_by = 'greencra' or closed_by = 'smitheli') and Board_Name = 'Labtech Alerts - Internal'")
    @infa_o = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening}  and Board_Name = 'Infrastructure Team' and (closed_by != 'ltadmin' or closed_by != 'Zadmin')")
    @comm_o = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'SC Comm' or team_name = 'SC Commercial')")
    @sc_o = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (team_name = 'SC Goverment' or team_name = 'SC Govt Internal') and (closed_by != 'Zadmin' and closed_by != 'ltadmin')")
    @nc_o = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'NC Govt' or team_name = 'NC Internal')")
    @fcr_o = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and team_name = 'FCR'")
    @ga_o = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (team_name = 'GA Govt' or team_name = 'GA Govt Internal') and (closed_by != 'Zadmin' and closed_by != 'ltadmin')")
    @lt_o = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and Board_Name = 'Labtech Alerts - Internal' and (closed_by != 'ltadmin' or closed_by != 'Zadmin')")
    @atrees = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and Detail_Description contains 'Page ID'")

@total_c = @fcr_c + @infa_c + @comm_c + @sc_c + @nc_c + @ga_c + @lt_c 
@total_o = @fcr_o  + @comm_o + @sc_o + @nc_o + @ga_o 

 hash = { :day => @morning, :total_c => @total_c, :fcr_c => @fcr_c, :infa_c => @infa_c, :comm_c => @comm_c, :sc_c => @sc_c, :nc_c => @nc_c, :ga_c => @ga_c, :lt_c => @lt_c, :total_o => @total_o, :fcr_o => @fcr_o, :infa_o => @infa_o, :comm_o => @comm_o, :sc_o => @sc_o, :nc_o => @nc_o, :ga_o => @ga_o, :lt_o => @lt_o,:atrees => @atrees}
array.push(hash)
end
puts array
class Array
  def to_csv(csv_filename="hash.csv")
    require 'csv'
    CSV.open(csv_filename, "wb") do |csv|
      csv << first.keys # adds the attributes name on the first line
      self.each do |hash|
        csv << hash.values
      end
    end
  end
end
array.to_csv("hash.csv")
