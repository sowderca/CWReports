require 'business_time'
require 'connect-stoopid'
require 'axlsx'
require 'mailgun'
require 'date'
require 'Time'


#set up mailgun client
mg_client = Mailgun::Client.new("key-70f741f0b10b167ad20b09f841151af8")
mb_obj = Mailgun::MessageBuilder.new 

# Function for deleting files
def remove_file(file)
    File.delete(file)
end

#set up vars
comps = Array.new
p = Axlsx::Package.new
wb = p.workbook:w
excel = []
daily_values = Hash.new
start_date = Date.parse("#{monday}")
end_date = Date.parse("#{today}")
days = start_date.business_dates_until(end_date)


days.each do | modify |
    comps << modify.iso8601
end

class Wkday 
    attr_ :date, :team, :opened, :closed, :appletrees
    
    def initialize(date, team, opened, closed, appletrees)
        @date = date
        @team = team
        @opened = opened
        @closed = closed
        @appletrees = appletrees 
    end

    def ==(other)
        self.class === other and
            other.date == @date and 
            other.team == @team and
            other.opened == @opened and
            other.closed == @closed and 
            other.appletrees = @appletrees 
    end
    
    alias eql? ==
    
    def hash
        @date.hash ^ @team.hash ^ @opened.hash ^ @closed.hash ^ @appletrees.hash
    end
end     
comps each do | daily |
    d = Wkday.new("#{daily}, #{total_c}, #{closed_FCR}, #{closed_infa}, #{closed_Comm}, #{closed_SC}, #{closed_NC}, #{closed_GA}, #{labtech_closed}, #{appletrees}" 
    
# Connectwise API calls
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
# End Connectwise API calls

