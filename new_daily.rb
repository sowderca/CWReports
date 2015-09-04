require 'connect-stoopid'
require 'date'
require 'mailgun'
require 'rest-client'
require 'axlsx'
require 'time'
require 'chronic'

# Connectwise API
cw_connect = ConnectStoopid::ReportingClient.new('connect.vc3.com', 'vc3', 'dattoCW', 'mushySh@pe83', :client_logging => false)

# Settin up query params
monday = Chronic.parse('monday this week')
friday = Chronic.parse('friday this week')

weekdays = monday.business_dates_until(friday)
comps = Array.new
weekdays.each do |modify|
    comps << modify.iso8601
end
# Methods 
def remove_file(file)
    File.delete(file)
end
# Classes 
class Engineer
    @@array = Array.new
    attr_accessor :username, :team
    def self.all_instances
        @@array
    end
    def initialize(username, team)
        @username = username
        @team     = team
        @@array << self
    end
end 

cameron = Engineer.
