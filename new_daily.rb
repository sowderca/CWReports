require 'connect-stoopid'
require 'date'
require 'mailgun'
require 'rest-client'
require 'axlsx'
require 'time'
require 'chronic'


# Connectwise API && Mailgun API
cw_connect = ConnectStoopid::ReportingClient.new('connect.vc3.com', 'vc3', 'dattoCW', 'mushySh@pe83', :client_logging => false)
cw_client = ConnectStoopid::ReportingClient.new('connect.vc3.com', 'vc3', 'dattoCW', 'mushySh@pe83')

# Excel and Email builder
mg_client = Mailgun::Client.new("key-70f741f0b10b167ad20b09f841151af8")
mb_obj = Mailgun::MessageBuilder.new


p  = Axlsx::Package.new
wb = p.workbook

excel = Array.new
excel << :tables


this_day          = Chronic.parse('today').iso8601
this_morning      = Chronic.parse('today at 7:00').iso8601
yesterday_evening = Chronic.parse('yesterday at 18:00').iso8601
yesterday_morning = Chronic.parse('yesterday at 7:00').iso8601
@this_morning = "[#{this_morning}]"
@yesterday_evening = "[#{yesterday_evening}]"
@yesterday_morning = "[#{yesterday_morning}]"

opened_Comm       = nil
opened_NC         = nil
opened_SC         = nil
opened_GA         = nil
opened_FCR        = nil
closed_Comm       = nil
closed_NC         = nil
closed_SC         = nil
closed_GA         = nil
closed_FCR        = nil
comm_closed       = nil
nc_closed         = nil
sc_closed         = nil
ga_closed         = nil
unresolved_comm   = nil
unresolved_nc     = nil
unresolved_sc     = nil
unresolved_ga     = nil
night_closed_GA   = nil
night_closed_NC   = nil
night_closed_SC   = nil
night_closed_Comm = nil
night_closed_FCR  = nil
night_opened_GA   = nil
night_opened_NC   = nil
night_opened_SC   = nil
night_opened_Comm = nil
night_opened_FCR  = nil


# FCR Shift
# -------------------------------------------------------------------------
loop do 
    opened_Comm = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_morning} and date_entered < #{@yesterday_evening} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'SC Comm' or team_name = 'SC Commercial')")
    break unless opened_Comm.nil? 
end   
loop do 
    opened_NC = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_morning} and date_entered < #{@yesterday_evening} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'NC Govt' or team_name = 'NC Internal')")
    break unless opened_NC.nil? 
end
loop do 
    opened_SC = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_morning} and date_entered < #{@yesterday_evening} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'SC Government' or team_name = 'SC Govt Internal')")
    break unless opened_SC.nil? 
end
loop do 
    opened_GA = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_morning} and date_entered < #{@yesterday_evening} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'GA Govt' or team_name = 'GA Govt Internal')")
    break unless opened_GA.nil? 
 end
loop do 
    opened_FCR = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_morning} and date_entered < #{@yesterday_evening} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'FCR')")
    break unless opened_FCR.nil? 
end
loop do 
    closed_Comm = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@yesterday_morning} and date_closed < #{@yesterday_evening} and (Board_Name = 'SC Comm Services' or Board_Name = 'SC Comm Internal') and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'washingt' or closed_by = 'browndea' or closed_by = 'baileypa' or closed_by = 'sleeperk')")
     break unless closed_Comm.nil? 
end
loop do 
    closed_NC = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@yesterday_morning} and date_closed < #{@yesterday_evening} and (Board_Name = 'NC Govt Services' or Board_Name = 'NC Govt Internal') and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'washingt' or closed_by = 'browndea' or closed_by = 'baileypa' or closed_by = 'sleeperk')")
    break unless closed_NC.nil? 
end
loop do 
    closed_SC = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@yesterday_morning} and date_closed < #{@yesterday_evening} and (Board_Name = 'SC Govt Services' or Board_Name = 'SC Govt Internal') and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'washingt' or closed_by = 'browndea' or closed_by = 'baileypa' or closed_by = 'sleeperk')")
    break unless closed_SC.nil? 
end
loop do 
    closed_GA = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@yesterday_morning} and date_closed < #{@yesterday_evening} and (Board_Name = 'GA Govt Services' or Board_Name = 'GA Govt Internal') and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'washingt' or closed_by = 'browndea' or closed_by = 'baileypa' or closed_by = 'sleeperk')")
    break unless closed_GA.nil? 
end
loop do 
    closed_FCR = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@yesterday_morning} and date_closed < #{@yesterday_evening} and Board_Name = 'FCR' and (closed_by = 'sowderca' or closed_by = 'carterma' or closed_by = 'washingt' or closed_by = 'browndea' or closed_by = 'baileypa' or closed_by = 'sleeperk')")
    break unless closed_FCR.nil? 
end


# Teams 
# --------------------------------------------------------------------------------
loop do 
     comm_closed= cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@yesterday_morning} and date_closed < #{@yesterday_evening} and (closed_by = 'blackmor' or closed_by = 'hutchinj' or closed_by = 'ricanora' or closed_by = 'tronconi' or closed_by = 'turnerbr' or closed_by = 'vanduzew') and (Board_Name = 'SC Comm Services' or Board_Name = 'SC Comm Internal')")
     break unless comm_closed.nil? 
end
loop do 
    nc_closed = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_morning} and date_entered < #{@yesterday_evening} and (closed_by = 'caywoodc' or closed_by = 'watsonro' or closed_by = 'pennellk' or closed_by = 'goodallp' or closed_by = 'offnertj' or closed_by = 'carterj') and (Board_Name = 'NC Govt Services' or Board_Name = 'NC Govt Internal')")
    break unless nc_closed.nil? 
end
loop do 
    sc_closed = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_morning} and date_entered < #{@yesterday_evening} and (closed_by = 'cgod' or closed_by = 'floydshe' or closed_by = 'gleatony' or closed_by = 'stetzpet' or closed_by = 'tuckerdu') and (Board_Name = 'SC Govt Services' or Board_Name = 'SC Govt Internal')")
    break unless sc_closed.nil? 
end
loop do 
    ga_closed = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_morning} and date_entered < #{@yesterday_evening} and (closed_by = 'burgettr' or closed_by = 'mcquiths' or closed_by = 'coxjosh' or closed_by = 'billupsd' or closed_by = 'paynesim' or closed_by = 'keddingl') and (Board_Name = 'GA Govt Services' or Board_Name = 'GA Govt Internal')")
    break unless ga_closed.nil? 
end
loop do 
    unresolved_comm = cw_client.run_report_count("reportName" => "Service", "conditions" => "(Closed_Flag = 'False' and status_description != 'Resolved' and status_description != 'Resolved (RMM)') and (Board_Name = 'FCR' or Board_Name = 'SC Comm Services' or Board_Name = 'SC Comm Internal')")
    break unless unresolved_comm.nil? 
end
loop do 
    unresolved_nc = cw_client.run_report_count("reportName" => "Service", "conditions" => "(Closed_Flag = 'False' and status_description != 'Resolved' and status_description != 'Resolved (RMM)') and (Board_Name = 'NC Govt Services' or Board_Name = 'NC Govt Internal')")
    break unless unresolved_nc.nil? 
end
loop do 
    unresolved_sc = cw_client.run_report_count("reportName" => "Service", "conditions" => "(Closed_Flag = 'False' and status_description != 'Resolved' and status_description != 'Resolved (RMM)') and (Board_Name = 'SC Govt Services' or Board_Name = 'SC Govt Internal')")
    break unless unresolved_sc.nil? 
end
loop do 
    unresolved_ga = cw_client.run_report_count("reportName" => "Service", "conditions" => "(Closed_Flag = 'False' and status_description != 'Resolved' and status_description != 'Resolved (RMM)') and (Board_Name = 'GA Govt Services' or Board_Name = 'GA Govt Internal')")
    break unless unresolved_ga.nil? 
end

# 2nd and 3rd Shift
# ------------------------------------------------------------------------------------------
loop do 
    night_opened_Comm = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_evening} and date_entered < #{@this_morning} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'SC Comm' or team_name = 'SC Commercial')")
    break unless night_opened_Comm.nil? 
end
loop do 
    night_opened_NC = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_evening} and date_entered < #{@this_morning} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'NC Govt' or team_name = 'NC Internal')")
    break unless night_opened_NC.nil? 
end
loop do 
    night_opened_SC = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_evening} and date_entered < #{@this_morning} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'SC Government' or team_name = 'SC Govt Internal')")
    break unless night_opened_SC.nil? 
end
loop do 
    night_opened_GA = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_evening} and date_entered < #{@this_morning} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'GA Govt' or team_name = 'GA Govt Internal')")
    break unless night_opened_GA.nil? 
end
loop do 
    night_opened_FCR = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_entered >= #{@yesterday_evening} and date_entered < #{@this_morning} and (closed_by != 'Zadmin' and closed_by != 'ltadmin') and (team_name = 'FCR')")
    break unless night_opened_FCR.nil? 
end
loop do 
    night_closed_Comm = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@yesterday_morning} and date_closed < #{@this_morning} and (Board_Name = 'SC Comm Services' or Board_Name = 'SC Comm Internal') and closed_by != 'Zadmin' and closed_by != 'ltadmin'")
    break unless night_closed_Comm.nil? 
end
loop do 
    night_closed_NC = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@yesterday_evening} and date_closed < #{@this_morning} and (Board_Name = 'NC Govt Services' or Board_Name = 'NC Govt Internal') and closed_by != 'Zadmin' and closed_by != 'ltadmin'")
    break unless night_closed_NC.nil? 
end
loop do 
    night_closed_SC = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@yesterday_evening} and date_closed < #{@this_morning} and (Board_Name = 'SC Govt Services' or Board_Name = 'SC Govt Internal') and closed_by != 'Zadmin' and closed_by != 'ltadmin'")
    break  unless night_closed_SC.nil? 
end
loop do 
    night_closed_GA = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@yesterday_evening} and date_closed < #{@this_morning} and (Board_Name = 'GA Govt Services' or Board_Name = 'GA Govt Internal') and closed_by != 'Zadmin' and closed_by != 'ltadmin'")
    break unless night_closed_GA.nil? 
end
loop do 
    night_closed_FCR = cw_client.run_report_count("reportName" => "Service", "conditions" => "date_closed >= #{@yesterday_evening} and date_closed < #{@this_morning} and Board_Name = 'FCR' and closed_by != 'Zadmin' and closed_by != 'ltadmin'") 
    break  unless night_closed_FCR.nil? 
end


# reg hours
# ------------------------------
daily_totalOpened = opened_Comm + opened_NC + opened_SC + opened_GA + opened_FCR rescue nil
daily_totalClosed = closed_Comm + closed_NC + closed_SC + closed_GA + closed_FCR rescue nil
fcr_rate = daily_totalClosed.to_f.send(:/, daily_totalOpened).send(:*,100) rescue nil
fcr_comm = closed_Comm.to_f.send(:/, opened_Comm).send(:*,100) rescue nil
fcr_nc = closed_NC.to_f.send(:/, opened_NC).send(:*,100) rescue nil
fcr_sc = closed_SC.to_f.send(:/, opened_SC).send(:*,100) rescue nil
fcr_ga = closed_GA.to_f.send(:/, opened_GA).send(:*,100) rescue nil
fcr_fcr = closed_FCR.to_f.send(:/, opened_FCR).send(:*,100) rescue nil


# night hours
# -----------------------------
nightly_totalOpened = night_opened_Comm + night_opened_NC + night_opened_SC + night_opened_GA + night_opened_FCR rescue nil
nightly_totalClosed = night_closed_Comm + night_closed_NC + night_closed_SC + night_closed_GA + night_closed_FCR rescue nil

# totals 
# ----------------------------
tickets_opened = daily_totalOpened + nightly_totalOpened rescue nil
tickets_closed = daily_totalClosed + nightly_totalClosed + comm_closed + nc_closed + sc_closed + ga_closed rescue nil
tickets_unresolved = unresolved_comm + unresolved_nc + unresolved_sc + unresolved_ga rescue nil


if excel.include?(:tables)
    wb.add_worksheet(:name => "Daily Report") do |sheet|
        sheet.add_row ["Vc3 Metrics"]
        sheet.add_row ["",""]
        sheet.add_row ["Total closed today", "#{tickets_closed}"]
        sheet.add_row ["Total opened today", "#{tickets_opened}"]
        sheet.add_row ["Total unresolved to date", "#{tickets_unresolved}"]
        sheet.add_row ["",""]
        sheet.add_row ["",""]
        sheet.add_row ["FCR"]
        sheet.add_row ["Team", "Total", "FCR", "SC Comm", "NC Govt", "SC Govt", "GA Govt"]
        sheet.add_row ["Closed by FCR", "#{daily_totalClosed}", "#{closed_FCR}", "#{closed_Comm}", "#{closed_NC}", "#{closed_SC}", "#{closed_GA}"]
        sheet.add_row ["Opened per team", "#{daily_totalOpened}", "#{opened_FCR}", "#{opened_Comm}", "#{opened_NC}", "#{opened_SC}", "#{opened_GA}"]
        sheet.add_row ["FCR percent closed", "#{fcr_rate}", "#{fcr_fcr}", "#{fcr_comm}", "#{fcr_nc}", "#{fcr_sc}", "#{fcr_ga}"]
        sheet.add_row ["",""]
        sheet.add_row ["Second and third Shift"]
        sheet.add_row ["Team", "Total", "FCR", "SC Comm", "NC Govt", "SC Govt", "GA Govt"]
        sheet.add_row ["Closed", "#{nightly_totalClosed}", "#{night_closed_FCR}", "#{night_closed_Comm}", "#{night_closed_NC}", "#{night_closed_SC}", "#{night_closed_GA}"]
        sheet.add_row ["Opened", "#{nightly_totalOpened}", "#{night_opened_FCR}", "#{night_opened_Comm}", "#{night_opened_NC}", "#{night_opened_SC}", "#{night_opened_GA}"]
        sheet.add_row ["",""]
        sheet.add_row ["Regional Teams"]
        sheet.add_row ["Team", "SC Comm", "NC Govt", "SC Govt", "GA Govt"]
        sheet.add_row ["Closed tickets by team", "#{comm_closed}", "#{nc_closed}", "#{sc_closed}", "#{ga_closed}"]
        sheet.add_row ["Total unresloved to date by team", "#{unresolved_comm}", "#{unresolved_nc}", "#{unresolved_sc}", "#{unresolved_ga}"]
   end
end

p.serialize('daily.xlsx')
#Email set up and send
# Define the from address.
mb_obj.set_from_address("reports@vc3.com", {"first"=>"Cameron", "last" => "Sowder"});
# Define a to recipient.
mb_obj.add_recipient(:to, "cameron.sowder@vc3.com", {"first" => "Cameron", "last" => "Sowder"});
# Define a cc recipient.
mb_obj.add_recipient(:to, "Amy.McKeown@vc3.com", {"first" => "Amy", "last" => "McKeown"});
mb_obj.add_recipient(:to, "Team-FCR@vc3.com", {"first" => "Mark", "last" => "Carter"});
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
mb_obj.add_attachment("daily.xlsx");


# Send your message through the client
mg_client.send_message("sandboxd511cc2c450b4b2d93f4bf5f4d385dae.mailgun.org", mb_obj)
File.delete('daily.xlsx')
