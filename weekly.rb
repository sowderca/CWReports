require 'business_time'
require 'connect-stoopid'
require 'axlsx'
require 'mailgun'
require 'date'
require 'time'
#require 'fastercsv'


#Mailgun 
mg_client = Mailgun::Client.new("key-70f741f0b10b167ad20b09f841151af8")
mb_obj = Mailgun::MessageBuilder.new

def remove_file(file)
  File.delete(file)
end
=begin vars
Set the dates for the report
=end
p = Axlsx::Package.new
wb = p.workbook

excel = []
excel << :tables
start_date = Date.parse("#{monday}")
end_date = Date.parse("#{today}")
comps = Array.new #array for new dates
#gets only the weekdays from the start date untill the end date.
days = start_date.business_dates_until(end_date)


days.each do | modify | 
  comps << modify.iso8601
end

comps.each do |day|
	@morning = "[#{day}T07:00:00-04:00]"
	@evening = "[#{day}T18:00:00-04:00]"
	cw_client = ConnectStoopid::ReportingClient.new('connect.vc3.com', 'vc3', 'dattoCW', 'mushySh@pe83')
	opened_total = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and  (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'SC Comm' or team_name = 'SC Commercial' or team_name = 'NC' or team_name = 'NC Internal' or team_name = 'GA Govt' or team_name = 'GA Govt Internal' or team_name = 'SC Government' or team_name = 'SC Govt Internal' or team_name = 'FCR')")

	opened_Comm = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'SC Comm' or team_name = 'SC Commercial')")
	opened_NC = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'NC Govt' or team_name = 'NC Internal')")
	opened_fcr = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and team_name = 'FCR'")
	opened_SC = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (team_name = 'SC Goverment' or team_name = 'SC Govt Internal') and (closed_by != 'Zadmin' and closed_by != 'ltadmin')")
	opened_GA = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (team_name = 'GA Govt' or team_name = 'GA Govt Internal') and (closed_by != 'Zadmin' and closed_by != 'ltadmin')")
	closed_secondshift = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@morning} and date_closed < #{@evening} and closed_by = 'vanscivb'")
	closed_total = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (Board_Name = 'FCR' or Board_Name = 'SC Comm Services' or Board_Name = 'SC Comm Internal' or Board_Name = 'NC Govt Services' or Board_Name = 'NC Govt Services' or Board_Name = 'NC Govt Internal' or Board_Name = 'SC Govt Services' or Board_Name = 'SC Govt Internal' or Board_Name = 'GA Govt Services' or Board_Name = 'GA Govt Internal' or Board_Name = 'Infastructure Team') and (closed_by != 'Zadmin' and closed_by != 'ltadmin')")
	closed_FCR = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@morning} and date_closed < #{@evening}  and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'sleeperk' or  closed_by = 'betivast' or closed_by = 'keddingl') and (Board_Name = 'FCR' or Board_Name = 'SC Comm Services' or Board_Name = 'SC Comm Internal' or Board_Name = 'NC Govt Services' or Board_Name = 'NC Govt Internal' or Board_Name = 'GA Govt Services' or Board_Name = 'GA Govt Internal' or Board_Name = 'SC Govt Services' or Board_Name = 'SC Govt Internal' or Board_Name = 'Infrastructure Team')")
	closed_Comm = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'keddingl') and (Board_Name = 'SC Comm Services' or Board_Name = 'SC Comm Internal')")
	closed_NC = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'keddingl') and (Board_Name = 'NC Govt Services' or Board_Name = 'NC Govt Internal')")
	closed_SC = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'keddingl') and (Board_Name = 'SC Govt Services' or Board_Name = 'SC Govt Internal')")
	closed_GA = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'keddingl') and (Board_Name = 'GA Govt Services' or Board_Name = 'GA Govt Internal')")
	closed_infa = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'keddingl' or closed_by = 'sleeperk') and Board_Name = 'Infrastructure Team'")
	closed_onFCR = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by= 'sowderca' or closed_by = 'carterma' or closed_by = 'washingt' or closed_by = 'keddingl' or closed_by = 'sleeperk') and Board_Name = 'FCR'")
	appletrees = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and Detail_Description contains 'Page ID'")
	wheelhouse_opened = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (Detail_Description contains 'WH1' or Detail_Description contains 'WH2' or Detail_Description contains 'WH3' or Detail_Description contains 'WH4' or Detail_Description contains 'WH5' or Detail_Description contains 'WH6' or Detail_Description contains 'WH7' or Detail_Description contains 'WH8' or Detail_Description contains 'WH9' or Detail_Description contains 'WH10' or Detail_Description contains 'WH11' or Detail_Description contains 'WH12' or Summary contains 'Perf - Memory % Bytes' or Summary contains 'Perf - Processor %' or Summary contains 'Perf -Total CPU Processor Time' or Summary contains 'AV - Disabled' or Summary contains 'AV - Missing' or Summary contains 'PF - 90% Plus Avg CPU' or Summary contains 'missing or stopped' or Summary contains 'IE Memory Usage' or Summary contains 'Print Jobs over Threshold')")
	wheelhouse_closed = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (Detail_Description contains 'WH1' or Detail_Description contains 'WH2' or Detail_Description contains 'WH3' or Detail_Description contains 'WH4' or Detail_Description contains 'WH5' or Detail_Description contains 'WH6' or Detail_Description contains 'WH7' or Detail_Description contains 'WH8' or Detail_Description contains 'WH9' or Detail_Description contains 'WH10' or Detail_Description contains 'WH11' or Detail_Description contains 'WH12' or Summary contains 'Perf - Memory % Bytes' or Summary contains 'Perf - Processor %' or Summary contains 'Perf -Total CPU Processor Time' or Summary contains 'AV - Disabled' or Summary contains 'AV - Missing' or Summary contains 'PF - 90% Plus Avg CPU' or Summary contains 'missing or stopped' or Summary contains 'IE Memory Usage' or Summary contains 'Print Jobs over Threshold') and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'betivast' or closed_by = 'keddingl' or closed_by = 'sleeperk' or closed_by = 'washingt')")
	closed_total = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (Board_Name = 'FCR' or Board_Name = 'SC Comm Services' or Board_Name = 'SC Comm Internal' or Board_Name = 'NC Govt Services' or  Board_Name = 'NC Govt Internal' or Board_Name = 'SC Govt Services' or Board_Name = 'SC Govt Internal' or Board_Name = 'GA Govt Services' or Board_Name = 'GA Govt Internal') and (closed_by != 'Zadmin' and closed_by != 'ltadmin')")

	unresolved = cw_client.run_report_count("reportName" => "Service", "conditions" => "(Closed_Flag = 'False' and status_description != 'Resolved' and status_description != 'Resolved (RMM)') and (Board_Name = 'FCR' or Board_Name = 'SC Comm Services' or Board_Name = 'SC Comm Internal' or Board_Name = 'NC Govt Services' or Board_Name = 'NC Govt Internal' or Board_Name = 'GA Govt Services' or Board_Name = 'GA Govt Internal' or Board_Name = 'SC Govt Services' or Board_Name = 'SC Govt Internal' or Board_Name = 'Labtech Alerts internal')")

	comm_closed= cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by = 'blackmor' or closed_by = 'hutchinj' or closed_by = 'ricanora' or closed_by = 'tronconi' or closed_by = 'turnerbr' or closed_by = 'vanduzew') and (Board_Name = 'SC Comm Services' or Board_Name = 'SC Comm Internal')")

	nc_closed = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by = 'caywoodc' or closed_by = 'molbycol' or closed_by = 'schwarzm' or closed_by = 'goodallp' or closed_by = 'offnertj' or closed_by = 'highsmiw') and (Board_Name = 'NC Govt Services' or Board_Name = 'NC Govt Internal')")

	sc_closed = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by = 'cgod' or closed_by = 'floydshe' or closed_by = 'masseyia' or closed_by = 'stetzpet' or closed_by = 'tuckerdu' or closed_by = 'watsonro') and (Board_Name = 'SC Govt Services' or Board_Name = 'SC Govt Internal')")

	ga_closed = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@morning} and date_entered < #{@evening} and (closed_by = 'burgettr' or closed_by = 'mcquiths' or closed_by = 'coxjosh' or closed_by = 'billupsd') and (Board_Name = 'GA Govt Services' or Board_Name = 'GA Govt Internal')")
	unresolved_comm = cw_client.run_report_count("reportName" => "Service", "conditions" => "(Closed_Flag = 'False' and status_description != 'Resolved' and status_description != 'Resolved (RMM)') and (Board_Name = 'FCR' or Board_Name = 'SC Comm Services' or Board_Name = 'SC Comm Internal')")
	unresolved_nc = cw_client.run_report_count("reportName" => "Service", "conditions" => "(Closed_Flag = 'False' and status_description != 'Resolved' and status_description != 'Resolved (RMM)') and (Board_Name = 'NC Govt Services' or Board_Name = 'NC Govt Internal')")
	unresolved_sc = cw_client.run_report_count("reportName" => "Service", "conditions" => "(Closed_Flag = 'False' and status_description != 'Resolved' and status_description != 'Resolved (RMM)') and (Board_Name = 'SC Govt Services' or Board_Name = 'SC Govt Internal')")
	unresolved_ga = cw_client.run_report_count("reportName" => "Service", "conditions" => "(Closed_Flag = 'False' and status_description != 'Resolved' and status_description != 'Resolved (RMM)') and (Board_Name = 'GA Govt Services' or Board_Name = 'GA Govt Internal')")
	fcr_rate = closed_FCR.to_f.send(:/, opened_total).send(:*,100)
	fcr_comm = closed_Comm.to_f.send(:/, opened_Comm).send(:*,100)
	fcr_nc = closed_NC.to_f.send(:/, opened_NC).send(:*,100)
	fcr_sc = closed_SC.to_f.send(:/, opened_SC).send(:*,100)
	fcr_ga = closed_GA.to_f.send(:/, opened_GA).send(:*,100)
	fcr_closedon = closed_onFCR.to_f.send(:/, opened_fcr).send(:*,100)
	wheelhouse_percent = wheelhouse_opened.to_f.send(:/, opened_total).send(:*,100)
	fcr_wheelhouse = wheelhouse_closed.to_f.send(:/, wheelhouse_opened).send(:*,100)

#excel sheet set up
if excel.include? :tables
  wb.add_worksheet(:name => "#{day}") do |sheet|
	sheet.add_row ["VC3 Metrics"]
	sheet.add_row ["",""]
	sheet.add_row ["Total closed today", "#{closed_total}"]
	sheet.add_row ["Total opened today", "#{opened_total}"]
	sheet.add_row ["Total unresolved to date", "#{unresolved}"]
	sheet.add_row ["",""]
	sheet.add_row ["",""]
	sheet.add_row ["FCR"]
	sheet.add_row ["Team", "Total", "FCR", "SC Comm", "NC Govt", "SC Govt", "GA Govt", "Infrastructure", "Second Shift"]
	sheet.add_row ["Closed today by FCR", "#{closed_FCR}", "#{closed_onFCR}", "#{closed_Comm}", "#{closed_NC}", "#{closed_SC}", "#{closed_GA}", "#{closed_infa}", "#{closed_secondshift}"]
	sheet.add_row ["Opened today per team", "#{opened_total}", "#{opened_fcr}", "#{opened_Comm}", "#{opened_NC}", "#{opened_SC}", "#{opened_GA}"]
	sheet.add_row ["FCR Percent Closed", "#{fcr_rate}%", "#{fcr_closedon}%", "#{fcr_comm}%", "#{fcr_nc}%", "#{fcr_sc}%", "#{fcr_ga}%"]	
	sheet.add_row ["",""]
	sheet.add_row ["Appletree tickets", "#{appletrees}"]
	sheet.add_row ["Wheelhouse tickets closed", "#{wheelhouse_closed}"]
	sheet.add_row ["Wheelhouse tickets opened", "#{wheelhouse_opened}"]
	sheet.add_row ["Percent of Wheelhouse tickets over total tickets", "#{wheelhouse_percent}%"]
	sheet.add_row ["Wheelhouse FCR rate", "#{fcr_wheelhouse}"]
	sheet.add_row ["",""]
	sheet.add_row ["",""]
	sheet.add_row ["Regional Teams"]
	sheet.add_row ["Team", "SC Comm", "NC Govt", "SC Govt", "GA Govt"]
	sheet.add_row ["Closed tickets today by team", "#{comm_closed}", "#{nc_closed}", "#{sc_closed}", "#{ga_closed}"]
	sheet.add_row ["Total unresolved to date by team", "#{unresolved_comm}", "#{unresolved_nc}", "#{unresolved_sc}", "#{unresolved_ga}"]
  end
end
end
puts start_date
p.serialize('Weekly.xlsx')
#Email set up and send
# Define the from address.
mb_obj.set_from_address("reports@vc3.com", {"first"=>"Cameron", "last" => "Sowder"});
# Define a to recipient.
mb_obj.add_recipient(:to, "cameron.sowder@vc3.com", {"first" => "Cameron", "last" => "Sowder"});
# Define a cc recipient.
mb_obj.add_recipient(:to, "Amy.McKeown@vc3.com", {"first" => "Amy", "last" => "McKeown"});
mb_obj.add_recipient(:to, "mark.carter@vc3.com", {"first" => "Mark", "last" => "Carter"});
# Define the subject.
mb_obj.set_subject("FCR Daily Report");
# Define the body of the message.
mb_obj.set_text_body("Daily Report");
# Set the Message-Id header. Pass in a valid Message-Id.
mb_obj.set_message_id("<2014101400 0000.11111.11111@example.com>")
# Clear the Message-Id header. Pass in nil or empty string.
mb_obj.set_message_id(nil)
mb_obj.set_message_id('')
# Other Optional Parameters.
mb_obj.add_attachment("Weekly.xlsx");


# Send your message through the client
mg_client.send_message("sandboxd511cc2c450b4b2d93f4bf5f4d385dae.mailgun.org", mb_obj)
remove_file("Weekly.xlsx")