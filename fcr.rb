require 'connect-stoopid'
require 'date'
require 'time'
require 'mailgun'
morning = (Date.today).strftime("%Y-%m-%d")
evening = Time.now.iso8601
@morning = "[#{morning}T07:00:00-04:00]"
@evening = "[#{evening}"

cw = ConnectStoopid::ReportingClient.new('connect.vc3.com', 'vc3', 'dattoCW', 'mushySh@pe83', :client_logging => false)
mg_client = Mailgun::Client.new("key-70f741f0b10b167ad20b09f841151af8")



@camerons = cw.run_report_count("reportName" => "Service", "conditions" => "closed_by = 'sowderca' and date_closed >= #{@morning}")
@markc = cw.run_report_count("reportName" => "Service", "conditions" => "closed_by  = 'carterma' and date_closed >= #{@morning}")
@tonyb = cw.run_report_count("reportName" => "Service", "conditions" => "closed_by = 'betivast' and date_closed >= #{@morning}")
@kevens = cw.run_report_count("reportName" => "Service", "conditions" => "closed_by  = 'sleeperk' and date_closed >= #{@morning}")
@theos = cw.run_report_count("reportName" => "Service", "conditions" => "closed_by  = 'washingt' and date_closed >= #{@morning}")
@bryans = cw.run_report_count("reportName" => "Service", "conditions" => "closed_by  = 'wigginsb' and date_closed >= #{@morning}")
@deantes = cw.run_report_count("reportName" => "Service", "conditions" => "closed_by = 'browndea' and date_closed >= #{@morning}")

message_params = {:from => "reports@vc3.com",
                  :to => "Team-FCR@vc3.com",
                  :subject => "Tickets closed per member",
                  :text => "cameron : #{@camerons}\
                            mark : #{@markc}\
                            tony : #{@tonyb}\
                            keven: #{@kevens}\
                            theo : #{@theos}\
                            bryan: #{@bryans}\
                            deante: #{@deantes}"}

mg_client.send_message "sandboxd511cc2c450b4b2d93f4bf5f4d385dae.mailgun.org", message_params
