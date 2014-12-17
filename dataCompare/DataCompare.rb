#*******************************************************************************************************************
#
#  Comparison Module
#
#*******************************************************************************************************************

begin
  require_relative "interface_module.rb"
  require 'csv'
  require 'win32ole'
rescue Exception => e
  #puts "Exception Occurred : " + e.class.to_s
end

#Module to compare data in excel and text files.
module DataCompare
#Compare class will compare two excel files and stores the results in the array arr_compare_results
  class Compare < CompareExcel
    #Compare Function to compare two excel files
    def compare_excel_files(col_name,col_pos,failed_col)
      arr_compare_result = []
      int_base_row_count = base_row_count
      int_current_row_count = current_row_count
      int_base_col_count =   base_column_count
      int_current_column_count = current_column_count
      if base_sheet_count == current_sheet_count
        arr_compare_result << "Pass !!!! The sheet counts of base and current versions are same."
      else
        arr_compare_result << "Warning !!!! The sheet counts of base and current versions are different."
      end
      if int_base_row_count == int_current_row_count
        arr_compare_result << "Pass !!!! The row counts of base: #{int_base_row_count} and current : #{int_current_row_count} versions are same."
      else
        arr_compare_result << "Warning !!!! The row counts of base: #{int_base_row_count} and current : #{int_current_row_count} versions are different."
      end

      flg_start_cmp = 0
      #Acquiring the Header Array for Base line file

      $base_header_arr,$int_start_base_data_line =  base_header_arr(col_name, col_pos)
      if $base_header_arr.empty?
        arr_compare_result << "Warning !!!! start column: #{col_name} is not present in the current version"
        flg_start_cmp = 1
      end

      #Acquiring the Header array for Current file
      $current_header_arr,$int_start_current_data_line =  current_header_arr(col_name, col_pos)
      if $current_header_arr.empty?
        arr_compare_result << "Warning !!!! start column: #{col_name} is not present in the current version"
        flg_start_cmp = 1
      end

      if int_base_col_count == int_current_column_count
        arr_compare_result << "Pass !!!! The column counts of base: #{int_base_col_count} and current : #{int_current_column_count} versions are same."
      else
        str_col_flag = 1
        int_col_diff = int_current_column_count - int_base_col_count
        arr_compare_result << "Warning !!!! The column counts of base :: #{int_base_col_count} and current : #{int_current_column_count} versions are different, Difference count: #{int_col_diff} and additional column names : #{$base_header_arr - $current_header_arr}"
      end

      a=1
      k=0
      #if flg_start_cmp == 0 then
      #  #***************************************************
      #  if not(failed_col.length==0 or failed_col.nil?)
      #    failed_col.each do |strPassArrycnt|
      #      $base_header_arr.each do |strPassArry|
      #        b=$int_start_base_data_line
      #        #puts "k is " +k.to_s
      #        if failed_col.length!=k
      #          if strPassArry.to_s.include? failed_col[k]
      #            l = $int_start_current_data_line
      #            m = $int_start_base_data_line
      #            arr_compare_result<<  "***!!!******************************Expected to Fail => column name compared is - #{strPassArry}*********************************************"
      #            for j in (b..int_base_row_count) do
      #              if  current_cell_data(l,a).include? base_cell_data(m,a)
      #                arr_compare_result<<  "Pass !!!! Before value #{base_cell_data(m,a)} match with the After value #{current_cell_data(l,a)} for the cell (#{j},#{a})"
      #              else
      #                arr_compare_result<<  "Warning !!!! Before value #{base_cell_data(m,a)} is not matching with the After value #{current_cell_data(l,a)} for the cell (#{j},#{a})"
      #              end
      #              b=b+1
      #              l=l+1
      #              m=m+1
      #            end
      #            k=k+1
      #          end
      #          a=a+1
      #        end
      #      end
      #    end
      #  end
      #  #****************************************************************
      $base_header_arr.each do |strPassArry|

        b = $int_start_base_data_line

        if strPassArry.to_s.include? $base_header_arr[k].to_s
          l = $int_start_current_data_line
          m = $int_start_base_data_line
          arr_compare_result<<  "***!!!**************************Column name compared is - #{strPassArry}**********************************************************"
          for j in (b..int_base_row_count) do
            if  current_cell_data(l,a).to_s == base_cell_data(m,a).to_s
              #puts "Pass : " + current_cell_data(l,a) + "Current Cell : #{l}, #{a}" + ":" + base_cell_data(m,a)  +  "Base Cell : #{m}, #{a}"
              arr_compare_result<<"Data Match : Pass !!!! Baseline value : #{base_cell_data(m,a)}|  Current value : #{current_cell_data(l,a)}| Cell (#{j},#{a})"
            else
              arr_compare_result<<"Data Mismatch : Warning !!!! Baseline value : #{base_cell_data(m,a)}|  Current value : #{current_cell_data(l,a)}| Cell (#{j},#{a})"
              #puts "Fail : " + current_cell_data(l,a) + "Current Cell : #{l}, #{a}" + ":" + base_cell_data(m,a)  +  "Base Cell : #{m}, #{a}"
            end
            b=b+1
            l=l+1
            m=m+1
          end
          k=k+1
        else
          arr_compare_result<<  "Data Mismatch : Warning !!!! Column name : #{strPassArry}, not found in Baseline"
          k=k+1
        end
        a=a+1
      end

      @excel.Quit
      WIN32OLE.ole_free(@excel)
      @excel = nil
      #Close Excel Process

      return arr_compare_result
    end
  end

#TextCompare class will compare two text files and stores the results in the array arr_compare_results
  class TextCompare < TextFile
    def compare_text(key_count)
      arr_before = []
      arr_after = []
      arr_results = []
      str_before_lines = @file_before.readlines
      str_before_lines.each do |line|
        if not line.chop.length == 0
          line = line.gsub("\n","")
          arr_before << line.split("||")
        end
      end
      @file_before.close
      str_after_lines = @file_after.readlines
      str_after_lines.each do |line|
        if not line.chop.length == 0
          line = line.gsub("\n","")
          arr_after << line.split("||")
        end
      end
      @file_after.close
      arr_len = arr_before[1].length
      puts Time.now
      if arr_before.length == arr_after.length
        puts arr_before.length
        puts arr_after.length
        arr_results << ["PASS | Baseline Record Count : #{arr_before.length}| Current Record Count : #{arr_after.length}"]
        for i in 1..arr_before.length-1
          for j in 1..arr_after.length-1
            flag = false
            if arr_before[i].take(key_count).eql?arr_after[j].take(key_count)
              flag = true
              flag_data_cmp = true
              $pass_data = ""
              $status = "Mismatches: "
              for l in key_count..arr_len-1
                if arr_before[i][l].strip.eql?arr_after[j][l].strip
                  $pass_data = "#{$pass_data}|#{arr_after[j][l]}"
                else
                  $pass_data = "#{$pass_data}|#{arr_before[i][l]}|#{arr_after[j][l]}|"
                  $status = "#{$status},#{arr_before[0][l].to_s}, "
                  flag_data_cmp = false
                end
              end
              if  flag_data_cmp
                arr_results <<["PASS | DATA MATCHED | KEY : #{arr_before[i].take(key_count).join("|")}|#{$pass_data}"]
              else
                arr_results << ["FAIL |DATA MISMATCHED |  KEY : #{arr_before[i].take(key_count).join("|")} |#{$status} #{$pass_data}"]
              end
              arr_after.delete_at(j)
              break
            end
            if i == arr_before.length/4
              puts "25% complete"
            elsif i == arr_before.length/2
              puts "50% complete"
            elsif i== arr_before.length/4 + arr_before.length/2
              puts "75% complete"
            end
          end
          if not flag
            arr_results << "FAIL | BASELINE DATA NOT FOUND | #{arr_before[i].join("|")}"
          end
        end
        if arr_after.length!= 0
          for i in 1.. arr_after.length-1
            arr_results << "FAIL | CURRENT DATA NOT FOUND | #{arr_after[i].join("|")}"
          end
        end
      elsif arr_before.length < arr_after.length
        arr_results << "FAIL | Baseline Record Count : #{arr_before.length}| Current Record Count : #{arr_after.length}"
        puts arr_before.length
        puts arr_after.length
        for i in 1..arr_after.length-1
          flag = false
          for j in 1..arr_before.length-1
            flag = false
            if arr_before[j].take(key_count).eql?arr_after[i].take(key_count)
              flag = true
              flag_data_cmp = true
              $status = "Mismatches: "
              $pass_data = ""
              for l in key_count..arr_len-1
                if arr_before[j][l].strip.eql?arr_after[i][l].strip
                  $pass_data = "#{$pass_data}| #{arr_after[i][l]}|"
                else
                  $pass_data = " #{$pass_data}|#{arr_before[j][l]} |#{arr_after[i][l]}|"
                  $status = "#{$status}#{arr_before[0][l].to_s}, "
                  flag_data_cmp = false
                end
              end
              if flag_data_cmp
                arr_results <<["PASS |DATA MATCHED| KEY : #{arr_after[i].take(key_count).join("|")} | #{$pass_data}"]
              else
                arr_results << ["FAIL |DATA MISMATCHED | KEY : #{arr_after[i].take(key_count).join("|")} |#{$status}#{$pass_data}"]
              end
              arr_before.delete_at(j)
              break
            end
          end
          if i == arr_before.length/4
            puts "25% complete"
          elsif i == arr_before.length/2
            puts "50% complete"
          elsif i== arr_before.length/4 + arr_before.length/2
            puts "75% complete"
          end
          if not flag
            arr_results <<[ "FAIL | CURRENT DATA NOT FOUND | #{arr_after[i].join("|")}"]
          end
        end
        if arr_before.length!=0
          for i in 1..arr_before.length-1
            arr_results << ["FAIL | BASELINE DATA NOT FOUND | #{arr_before[i].join("|")}"]
          end
        end
      else
        arr_results << "FAIL | Baseline Record Count : #{arr_before.length}| Current Record Count : #{arr_after.length}"
        puts arr_before.length
        puts arr_after.length
        for i in 1..arr_before.length-1
          flag = false
          for j in 1..arr_after.length-1
            flag = false
            if arr_before[i].take(key_count).eql?arr_after[j].take(key_count)
              flag = true
              flag_data_cmp = true
              $pass_data = ""
              $status = "Mismatches: "
              for l in key_count..arr_len-1
                if arr_before[i][l].strip.eql?arr_after[j][l].strip
                  $pass_data = "#{$pass_data}|#{arr_after[j][l]}"
                else
                  $pass_data = "#{$pass_data}|#{arr_before[i][l]}|#{arr_after[j][l]}|"
                  $status = "#{$status},#{arr_before[0][l].to_s}, "
                  flag_data_cmp = false
                end
              end
              if  flag_data_cmp
                arr_results <<["PASS |DATA MATCHED | KEY : #{arr_before[i].take(key_count).join("|")}|#{$pass_data}"]
              else
                arr_results << ["FAIL |DATA MISMATCHED | KEY : #{arr_before[i].take(key_count).join("|")} |#{$status} #{$pass_data}"]
              end
              arr_after.delete_at(j)
              break
            end
          end
          if i == arr_before.length/4
            puts "25% complete"
          elsif i == arr_before.length/2
            puts "50% complete"
          elsif i== arr_before.length/4 + arr_before.length/2
            puts "75% complete"
          end
          if not flag
            arr_results << "FAIL | BASELINE DATA NOT FOUND | #{arr_before[i].join("|")}"
          end
        end
        if arr_after.length!=0
          for i in 1.. arr_after.length-1
            arr_results << "FAIL | CURRENT DATA NOT FOUND | #{arr_after[i].join("|")}"
          end
        end
      end

      return arr_results
    end
    def compare_sql_text(before_path,after_path)
      @file_before = File.open(before_path)
      @file_after = File.open(after_path)
      str_result = []
      str_before_data = @file_before.readlines
      str_after_data = @file_after.readlines

      if str_before_data == str_after_data
        str_result << "Pass !!!! The text in the files match"
      else
        str_result << "Warning !!!! The comparison mismatched "
        str_result << "Warning !!!! Before data : #{str_before_data} After Data : #{str_after_data}"
      end
      @file_before.close
      @file_after.close
      return str_result
    end
  end

  #Method results_csv takes in the arr_compare_results and writes it in a cvs file
  def results_csv(strResultsCSVPath, strResults,pass_req)

    if File.exist?(strResultsCSVPath)
      File.delete(strResultsCSVPath)
    end
    File.open(strResultsCSVPath,"a")   do |f|
      strResults.each do |row|
        f.puts row
      end
    end
  end
end

