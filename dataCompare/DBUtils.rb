require 'rubygems'
require 'dbi'
class ConnectDB

  # Constructor to initialize the path of the Database
  def initialize(connectionName)
    @connect = connectionName
  end
  #Method to connect to the DB using DBI module

  def connect_db
    begin
      @dbh = DBI.connect(@connect)
      puts "Connection to the " + @connect.to_s + "  was successful"
    rescue Exception => e
      puts "An error occurred"
      puts "Error message: #{e.class}"
    end
  end

  def connect_icm_db(user, pwd)
    begin
      @dbh = DBI.connect(@connect,user,pwd)
      puts "Connection to: " + @connect.to_s + " is successful"
    rescue Exception => e
      puts "Exception Raised : " + e.class.to_s
    end
  end

  # fetch all the rows & columns data from the sql query provided
  def give_me_all_rows(sql)
    db_row = []
    @sth =  @dbh.execute(sql)
    db_row= @sth.map{ |row| row.to_a }
    return db_row
  end

  def give_col_names(sql)
    db_col = []
    @sth =  @dbh.execute(sql)
    db_col = @sth.column_names
    return db_col
  end

  def fetch_rows_text(sql, str_file_path)

    col =  give_col_names(sql)
    rows = give_me_all_rows(sql)
    File.open(str_file_path,"w") do |f|
      f << col
      f.puts "\n"
      rows.each do |row|
        row.each do |j|
          f << "#{j}||"
        end
        f.puts "\n"
      end
      f.close
    end

  end

  # Method to count no. of records returned by the query
  def count_records(sql)
    sth = @dbh.execute(sql)
    rows = sth.fetch_all()
    sth.finish
    return rows.count
  end

  # Method to execute SQL query using Execute statement
  def execute(sql)
    #@connection.Execute(sql)
    @dbh.execute(sql)
  end

  #Method to close the connection to the db
  def close
    @dbh.disconnect
  end
end

