require 'sequel'
require 'csv'
Sequel::Model.plugin :csv_serializer
DB = Sequel.sqlite
DB.create_table :items do 
    primary_key :#blala
    String      :#blahaha
    Float       :#blalala
dataset = DB[:items]


csv_text = File.read('/home/sowderca/Desktop/hash.csv')
csv = CSV.parse(csv_text, :headers => true)
csv.each do |row|
    Moulding.create!(row.to_hash)
end
