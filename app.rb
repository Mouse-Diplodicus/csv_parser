require 'csv'

HEADER_HASH = Hash.new

def main(argv = [])

  puts "Starting csv cleaner..."

  if (argv.empty?)
    puts "ERROR: No root directory was specified."
    puts "You must pass in a path as the starting directory."
    exit
  else
    root_dir = File.expand_path(argv.shift)
  end
  puts "root directory=#{root_dir}"

  output_path = calc_output_path(root_dir)
  puts "output path=#{output_path}"

  csv_dir = "#{root_dir}/Example Files/eplusmtr.csv"
  csv_table = load_csv(csv_dir)
  parse_csv(csv_table, output_path)
end

def load_csv(dir)
  data = CSV.read(dir, headers: true)
  return data
end

def parse_csv(csv_table, output_path)
  build_header_hash(csv_table)
  list_to_remove = query_user_for_headers()
  puts "Removing columns selected by user"
  csv_table.by_col!()
  for key in list_to_remove
    csv_table.delete(HEADER_HASH[key])
    puts("deleted column: #{HEADER_HASH[key]}")
  end
  csv_table.by_row!()
  output_csv = CSV.open(output_path, "a+", headers: true)
  if(output_csv.header_row?())
    output_csv << csv_table.headers()
  end
  puts "Removing non-hourly data points"
  is_hourly_data = false
  for row in csv_table
    if !is_hourly_data and row['Date/Time'][1,5] == "01/01"
      is_hourly_data = true
    end
    if is_hourly_data and row['Date/Time'][11,2] == "00"
      output_csv << row
    end
  end
end

def calc_output_path(root_dir)
  t = Time.new
  t_string = "#{t.month}-#{t.day}_#{t.hour}h_#{t.min}m_#{t.sec}s"
  output_path = "#{root_dir}/eplusmtr_clean_#{t_string}.csv"
  return output_path
end

def build_header_hash(csv_table)
  index = 0
  for header in csv_table.headers
    HEADER_HASH[index] = header
    index += 1
  end
end

def query_user_for_headers()
  HEADER_HASH.each do |key, header|
    if key != 0
      puts "#{key}:  #{header}"
    end
  end
  puts "Select which columns you would like to remove. Valid inputs include "
  puts "comma separated lists e.g. (1, 2, 5) or for a range input the starting"
  puts "and ending values seperated by a dash (-) e.g. (7-11)"
  user_input = gets.chomp
  user_input = user_input.split(/,/)
  list_to_remove = []
  for item in user_input
    if !item.include? '-'
      list_to_remove << item.to_i
    else
      temp = item.split(/-/)
      temp[0] = temp[0].to_i
      temp[1] = temp[1].to_i
      while temp[0] <= temp[1]
        list_to_remove << temp[0]
        temp[0] += 1
      end
    end
  end
  puts "Application will remove the following columns from the output csv:"
  for key in list_to_remove
    puts "#{key} - #{HEADER_HASH[key]}"
  end
  invalid_input = true
  while invalid_input
    puts "press y to confirm, n to reselect columns, or q to cancel"
    user_input = gets.chomp
    case user_input.downcase
    when 'y'
      invalid_input = false
    when 'n'
      list_to_remove = query_user_for_headers()
      invalid_input = false
    when 'q'
      puts "quitting"
      exit(0)
    else
      puts "invalid input"
    end
  end
  return list_to_remove
end

main(ARGV)
