require 'sequel'
require 'sqlite3'
require 'mysql2'

DB = Sequel.sqlite # memory database, requires sqlite3
DB = Sequel.connect(:adapter => 'mysql2', :user => 'root', :host => 'localhost', :database => 'scanty',:password=>'l0peZ!1972')
DB.create_table :items do
  primary_key :id
  String :name
  Float :price
end