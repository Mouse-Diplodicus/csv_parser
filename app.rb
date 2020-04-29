require 'csv'
require 'json'
require 'spreadsheet'

HEADER_HASH = Hash.new
PAT_DIR_HASH = Hash.new
HEADERS = ['Date/Time', 'InteriorEquipment:Electricity [kWh](Hourly)',
   'InteriorEquipment:Gas [therm](Hourly)', 'InteriorLights:Electricity [kWh](Hourly)',
   'ExteriorLights:Electricity [kWh](Hourly)', 'WaterSystems:Electricity [kWh](Hourly)',
   'WaterSystems:Gas [therm](Hourly)',  'Pumps:Electricity [kWh](Hourly)',
   'Fans:Electricity [kWh](Hourly)',  'Heating:Electricity [kWh](Hourly)',
   'Heating:Gas [therm](Hourly)',  'Cooling:Electricity [kWh](Hourly)',
   'Whole Building:Facility Total Electric Demand Power [kWh](Hourly)',
   'Electricity:Facility [kWh](Hourly)', 'Gas:Facility [therm](Hourly)']

HEADER_CONVERT_HASH = Hash['InteriorEquipment:Electricity [J](Hourly)' => 'InteriorEquipment:Electricity [kWh](Hourly)',
                           'InteriorEquipment:Gas [J](Hourly)' => 'InteriorEquipment:Gas [therm](Hourly)',
                           'InteriorLights:Electricity [J](Hourly)' => 'InteriorLights:Electricity [kWh](Hourly)',
                           'ExteriorLights:Electricity [J](Hourly)' => 'ExteriorLights:Electricity [kWh](Hourly)',
                           'WaterSystems:Electricity [J](Hourly)' => 'WaterSystems:Electricity [kWh](Hourly)',
                           'WaterSystems:Gas [J](Hourly)' => 'WaterSystems:Gas [therm](Hourly)',
                           'Pumps:Electricity [J](Hourly)' => 'Pumps:Electricity [kWh](Hourly)',
                           'Fans:Electricity [J](Hourly)' => 'Fans:Electricity [kWh](Hourly)',
                           'Heating:Electricity [J](Hourly)' => 'Heating:Electricity [kWh](Hourly)',
                           'Heating:Gas [J](Hourly)' => 'Heating:Gas [therm](Hourly)',
                           'Cooling:Electricity [J](Hourly)' => 'Cooling:Electricity [kWh](Hourly)',
                           'Whole Building:Facility Total Electric Demand Power [W](Hourly)' => 'Whole Building:Facility Total Electric Demand Power [kWh](Hourly)',
                           'Electricity:Facility [J](TimeStep)' => 'Electricity:Facility [kWh](Hourly)',
                           'Gas:Facility [J](TimeStep) ' => 'Gas:Facility [therm](Hourly)']

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
  puts "Output path=#{output_path}"

  build_pat_dir_hash(root_dir)
  workbook = Spreadsheet::Workbook.new
  PAT_DIR_HASH.each do |name, input_csv_dir|
    puts "Generating sheet for #{name}"
    sheet = workbook.create_worksheet :name => name

    csv_table = load_csv(input_csv_dir)
    csv_table = clean_csv(csv_table)
    csv_table.by_col!()

    HEADERS.each_with_index do |column_header, column|
      sheet[0, column] = column_header
      if HEADER_CONVERT_HASH.has_value? column_header
        original_header = HEADER_CONVERT_HASH.key(column_header)
        if column_header.include? '[kWh]' and original_header.include? '[J]'
          csv_table[original_header].each_with_index do |cell, row|
            sheet[row+1, column] = joules_to_kWh(cell.to_f)
          end
        elsif column_header.include? 'therm' and original_header.include? '[J]'
          csv_table[original_header].each_with_index do |cell, row|
            sheet[row+1, column] = joules_to_therm(cell.to_f)
          end
        elsif column_header.include? '[kWh]' and original_header.include? '[W]'
          csv_table[original_header].each_with_index do |cell, row|
            sheet[row+1, column] = cell.to_f/1000
          end
        end
      elsif column_header == 'Date/Time'
        csv_table[column_header].each_with_index do |cell, row|
          sheet[row+1, column] = cell
        end
      else
        csv_table[column_header].each_with_index do |cell, row|
          sheet[row+1, column] = cell.to_f
        end
      end
    end
    puts "Done"
  end
  puts 'Writing output file'
  workbook.write output_path
end

def build_pat_dir_hash(root_dir)
  file = File.read("#{root_dir}/pat.json")
  json = JSON.parse(file)
  data_points = json["datapoints"]
  data_points.each do |data_point|
    PAT_DIR_HASH[data_point["name"]] = "#{root_dir}/perm_data/analysis_#{data_point["analysis_id"]}/data_point_#{data_point["_id"]}/run/eplusout.csv"
  end
end

def load_csv(dir)
  data = CSV.read(dir, headers: true)
  return data
end

def clean_csv(csv_table)
  puts "Removing Design Days"
  design_days_removed = false
  count = 0
  while !design_days_removed
    csv_table.delete(1)
    count += 1
    if csv_table[1]['Date/Time'][1,5] == "01/01"
      design_days_removed = true
      puts "Design Days Removed: #{count}"
    end
  end

  puts "Removing non-hourly data points"
  csv_table.delete_if {|row| row['Date/Time'][11,2] != "00"}

  return csv_table
end

def calc_output_path(root_dir)
  t = Time.new
  t_string = "#{t.month}-#{t.day}_#{t.hour}h_#{t.min}m_#{t.sec}s"
  output_path = "#{root_dir}/eplusout_clean_#{t_string}.xls"
  return output_path
end

def joules_to_therm(joules)
  therm = joules/105480400
  return therm
end

def joules_to_kWh(joules)
  kWh = joules/3600000
  return kWh
end

main(ARGV)
