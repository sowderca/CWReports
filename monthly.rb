#gems needed
require 'business_time'
require 'connect-stoopid'
require 'axlsx'
require 'mailgun'
require 'date'
require 'time'
require 'pstore'


=begin 

####### Methods ############

=end
def remove_file(file)
    File.delete(file)
end 

=begin

###### Classes ############

=end
class Array
    def to_csv(csv_filename="month.csv")
        CSV.open(csv_filename, "wb") do |csv|
            csv << first.keys # adds the attributes name on the first line
            self.each do |hash|
                csv << hash.values
            end
        end
    end
end

###### vars ##############
total_fcr_hrcost = 105.99
total_monthlyhrs = 1213
total_fcrcost    = 128601.20
today = Date.today
month = Date.today.month
year  = Date.today.year
first_ofmonth = Date.parse("#{year}-#{month}-1")

days = first_ofmonth.business_dates_until(today)
days_iso = Array.new
array = Array.new
days.each do |modify|
    days_iso << modify.iso8601
end
days_iso.each do |day|
    @morning = "[#{day}T07:00:00-04:00]"
    @evening = "[#{day}T18:00:00-04:00]"

    cw_client = ConnectStoopid::ReportingClient.new('connect.vc3.com', 'vc3', 'dattoCW', 'mushySh@pe83', :client_logging => false)    
end
