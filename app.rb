require 'csv'

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

  csv_dir = "#{root_dir}/Example Files/eplusmtr.csv"
  output_path = calc_output_path(root_dir)
  puts "output path=#{output_path}"

  csv_input = load_csv(csv_dir)
  parse_csv(csv_input, output_path)
end

def load_csv(dir)
  data = CSV.read(dir, headers: true, header_converters: :symbol)
  return data
end

def parse_csv(input_csv, output_path)
  output_csv = CSV.open(output_path, "a+")
  puts input_csv.headers
  for row in input_csv
    if row[:datetime][11,2] == "00"
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

main(ARGV)
