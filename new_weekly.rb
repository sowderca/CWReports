require 'connect-stoopid'
require 'date'
require 'axlsx'
require 'chronic'
require 'mailgun'



cw_connect = ConnectStoopid::ReportingClient.new('connect.vc3.com', 'vc3', 'dattoCW', 'mushySh@pe83', :client_logging => false)
cw_client  = ConnectStoopid::ReportingClient.new('connect.vc3.com', 'vc3', 'dattoCW', 'mushySh@pe83')


p  = Axlsx::Package.new
wb = p.workbook

excel = Array.new
excel << :tables

monday    = Chronic.parse('monday this week').iso8601
tuesday   = Chronic.parse('tuesday this week').iso8601
wednesday = Chronic.parse('wednesday this week').iso8601
thursday  = Chronic.parse('thursday this week').iso8601
friday    = Chronic.parse('friday this week').iso8601


# Time tables | these times are defined by business time


module Inject
    def inject(n)
        each do |value|
            n = yield(n, value)
        end
        n
    end
    def sum(initial = 0)
        inject(initial) { |n, value| n + value}
    end
end


