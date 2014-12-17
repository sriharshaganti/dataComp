#*******************************************************************************************************************
#
#  Interface Module
#
#*******************************************************************************************************************


begin
  gem 'ruby-ole', '1.2.11.4'
  require 'win32ole'
rescue Exception => e
  #puts "Exception occurred"
end

#class to extract the required data from Excel file like sheet, row count, column count, cell values etc.....
class CompareExcel

  def initialize(str_before_path, str_after_path, sheet = nil)
   @excel = WIN32OLE.new("Excel.Application")
   @base_wk = @excel.WorkBooks.Open(str_before_path)
   @current_wk = @excel.WorkBooks.Open(str_after_path)
   @base_sheet = @base_wk.WorkSheets(sheet)
   @current_sheet = @current_wk.WorkSheets(sheet)
  end

  def base_sheet_count
   @base_wk.WorkSheets.count
  end

  def current_sheet_count
   @current_wk.WorkSheets.count
  end

  def base_row_count
   @base_sheet.UsedRange.Rows.Count
  end

  def current_row_count
  @current_sheet.UsedRange.Rows.Count
  end

  def current_column_count
   @current_sheet.UsedRange.Columns.Count
  end

  def base_column_count
   @base_sheet.UsedRange.Columns.Count
  end

  def base_cell_data(row, col)
  @base_sheet.cells(row,col).value.to_s
  end

  def current_cell_data(row, col)
  @current_sheet.cells(row, col).value.to_s
  end

  def base_rows(row)
  @base_sheet.Rows(row).value.to_s
  end

  def current_rows(row)
   @current_sheet.Rows(row).value.to_s
  end

  def base_header_arr(col_name,col_pos)
   $arr_head_base = []
   i = 1
   j = 1
   int_base_row_count = base_row_count
   int_base_col_count = base_column_count

   if col_pos != ""
    for i in (1..int_base_row_count )
      #puts  base_cell_data(i,col_pos)
       if  base_cell_data(i,col_pos) == col_name.to_s
         int_start_line = i
         $int_start_data_line = i+1
         $arr_head_base = base_rows(int_start_line)
         break
       end
    end
   else
     for i in (1..int_base_row_count )
       for j in (1..int_base_col_count)
         if  base_cell_data(i,j) == col_name.to_s then
           int_start_line = i
           $int_start_data_line = i+1
           $arr_head_base = base_rows(int_start_line)
           break
         end
       end
     end
   end
  if $arr_head_base.length != 0
   $arr_head_base = $arr_head_base.split(",")
   $base_header_arr = []
   $arr_head_base.each do |i|
     if i.to_s.include? "nil"
     else
       $base_header_arr << i.to_s.gsub("[","")
     end
   end
   return $base_header_arr,$int_start_data_line
  end
  end

  def current_header_arr(col_name, col_pos)
   $arr_head_current = []
   i = 1
   j = 1
   int_current_row_count =  current_row_count
   int_current_column_count = current_column_count
   if col_pos != ""
    for i in (1..int_current_row_count )
     if  current_cell_data(i,col_pos) == col_name.to_s
       int_start_line = i
       $int_start_data_line = i+1
       $arr_head_current = current_rows(int_start_line)
       break
     end
    end
    else
      for i in (1..int_current_row_count )
        for j in (1..int_current_column_count)
        if  current_cell_data(i,j) == col_name.to_s  then
          int_start_line = i
          $int_start_data_line = i+1
          $arr_head_current = current_rows(int_start_line)
          break
        end
        end
      end
    end
  if $arr_head_current.length != 0
   $arr_head_current  =  $arr_head_current.split(",")
   $current_header_arr = []

   $arr_head_current.each do |i|
     if i.to_s.include? "nil"
     else
       $current_header_arr << i.to_s.gsub("[","")
     end
   end
   return $current_header_arr,$int_start_data_line
  end
  end

  def close_excel
    wmi = WIN32OLE.connect("winmgmts://")
    processes = wmi.ExecQuery("select * from win32_process where commandline like '%excel.exe\"% /automation %'")
    for process in processes do
      Process.kill( 'KILL', process.ProcessID.to_i)
    end
  end

end

#class  Open the  text files
class TextFile
  def initialize(str_before_path, str_after_path)
    @file_before = File.open(str_before_path)
    @file_after = File.open(str_after_path)
  end
  def close
    @file_before.close
    @file_after.close
  end
end