require 'csv'
require 'json'
require 'logger'
require 'spreadsheet'

LOGGER = Logger.new(STDOUT)
LOGGER.level = Logger::INFO

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

HEADER_CONVERT_HASH = Hash['InteriorEquipment:Electricity [kWh](Hourly)' => 'InteriorEquipment:Electricity [J](Hourly)',
                           'InteriorEquipment:Gas [therm](Hourly)' => 'InteriorEquipment:Gas [J](Hourly)',
                           'InteriorLights:Electricity [kWh](Hourly)' => 'InteriorLights:Electricity [J](Hourly)',
                           'ExteriorLights:Electricity [kWh](Hourly)' => 'ExteriorLights:Electricity [J](Hourly)',
                           'WaterSystems:Electricity [kWh](Hourly)' => 'WaterSystems:Electricity [J](Hourly)',
                           'WaterSystems:Gas [therm](Hourly)' => 'WaterSystems:Gas [J](Hourly)',
                           'Pumps:Electricity [kWh](Hourly)' => 'Pumps:Electricity [J](Hourly)',
                           'Fans:Electricity [kWh](Hourly)' => 'Fans:Electricity [J](Hourly)',
                           'Heating:Electricity [kWh](Hourly)' => 'Heating:Electricity [J](Hourly)',
                           'Heating:Gas [therm](Hourly)' => 'Heating:Gas [J](Hourly)',
                           'Cooling:Electricity [kWh](Hourly)' => 'Cooling:Electricity [J](Hourly)',
                           'Whole Building:Facility Total Electric Demand Power [kWh](Hourly)' => 'Whole Building:Facility Total Electric Demand Power [W](Hourly)',
                           'Electricity:Facility [kWh](Hourly)' => 'Electricity:Facility [J](TimeStep) ',
                           'Gas:Facility [therm](Hourly)' => 'Gas:Facility [J](TimeStep) ']

def main(argv = [])

  LOGGER.info("Starting csv cleaner")

  if (argv.empty?)
    LOGGER.error("ERROR: No root directory was specified.")
    LOGGER.warn("You must pass in a path as the starting directory.")
    exit
  else
    root_dir = File.expand_path(argv.shift)
  end
  LOGGER.info("root directory=#{root_dir}")

  output_path = calc_output_path(root_dir)
  LOGGER.info("Output path=#{output_path}")

  build_pat_dir_hash(root_dir)
  workbook = build_workbook()
  LOGGER.info('Writing output file')
  workbook.write output_path
  LOGGER.info('Done!')
end

def build_workbook()
  LOGGER.info("Starting to build workbook")
  workbook = Spreadsheet::Workbook.new
  PAT_DIR_HASH.each do |name, input_csv_dir|
    LOGGER.info("Converting CSV table to sheet for #{name}")
    csv_table = load_csv(input_csv_dir)
    csv_table = clean_csv(csv_table)
    sheet = workbook.create_worksheet :name => name
    if name.include? "_HP"
      hp_conversion_hash = HEADER_CONVERT_HASH.clone
      hp_conversion_hash['WaterSystems:Electricity [kWh](Hourly)'] = 'HPWH TANK:Water Heater Heating Energy [J](Hourly)'
      hp_conversion_hash['Heating:Electricity [kWh](Hourly)'] = 'HPWH TANK 1:Water Heater Electric Energy [J](Hourly)'
      table_to_sheet(csv_table, sheet, hp_conversion_hash)
    elsif
      table_to_sheet(csv_table, sheet, HEADER_CONVERT_HASH)
    end
    LOGGER.info("Done")
  end
  return workbook
end

def table_to_sheet(csv_table, sheet, conversion_hash)

  csv_table.by_col!()
  HEADERS.each_with_index do |column_header, column|
    sheet[0, column] = column_header
    if conversion_hash.key?(column_header)
      original_header = conversion_hash[column_header]
      LOGGER.debug("Generating #{column_header} column from #{original_header}")
      if column_header.include? '[kWh]' and original_header.include? '[J]'
        LOGGER.debug("Converting J to kWh")
        csv_table[original_header].each_with_index do |cell, row|
          sheet[row+1, column] = joules_to_kWh(cell.to_f)
        end
      elsif column_header.include? 'therm' and original_header.include? '[J]'
        LOGGER.debug("Converting J to therm")
        csv_table[original_header].each_with_index do |cell, row|
          sheet[row+1, column] = joules_to_therm(cell.to_f)
        end
      elsif column_header.include? '[kWh]' and original_header.include? '[W]'
        LOGGER.debug("Converting W to kWh")
        csv_table[original_header].each_with_index do |cell, row|
          sheet[row+1, column] = cell.to_f/1000
        end
      end
    elsif column_header == 'Date/Time'
      LOGGER.debug("Generating date/time column")
      csv_table[column_header].each_with_index do |cell, row|
        sheet[row+1, column] = cell
      end
    else
      csv_table[column_header].each_with_index do |cell, row|
        sheet[row+1, column] = cell.to_f
      end
    end
  end
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
  LOGGER.debug("Removing Design Days")
  design_days_removed = false
  count = 0
  while !design_days_removed
    csv_table.delete(1)
    count += 1
    if csv_table[1]['Date/Time'][1,5] == "01/01"
      design_days_removed = true
      LOGGER.debug("Design day datapoints removed: #{count}")
    end
  end

  LOGGER.debug("Removing non-hourly data points")
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
