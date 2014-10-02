#!/usr/bin/env ruby

require 'optparse'
require 'google/api_client'
require 'google/api_client/client_secrets'
require 'google/api_client/auth/file_storage'
require 'google/api_client/auth/installed_app'
require 'active_support/all'
require 'spreadsheet'

# set your application name
APPLICATION_NAME    = 'google-drive-permission-search'
APPLICATION_VERSION = '1.0.0'

API_VERSION           = 'v2'
CACHED_API_FILE       = "drive-#{API_VERSION}.cache"
CREDENTIAL_STORE_FILE = "#{$0}-oauth2.json"

debug = false

# #setup is originated by https://github.com/google/google-api-ruby-client-samples/tree/master/drive
def setup()
  client = Google::APIClient.new(:application_name    => APPLICATION_NAME,
                                 :application_version => APPLICATION_VERSION)

  file_storage = nil
  begin
    file_storage = Google::APIClient::FileStorage.new(CREDENTIAL_STORE_FILE)
  rescue URI::InvalidURIError
    File.delete CREDENTIAL_STORE_FILE
    file_storage = Google::APIClient::FileStorage.new(CREDENTIAL_STORE_FILE)
  end
  if file_storage.authorization.nil?
    client_secrets       = Google::APIClient::ClientSecrets.load
    flow                 = Google::APIClient::InstalledAppFlow.new(
      :client_id     => client_secrets.client_id,
      :client_secret => client_secrets.client_secret,
      :scope         => ['https://www.googleapis.com/auth/drive']
    )
    client.authorization = flow.authorize(file_storage)
  else
    client.authorization = file_storage.authorization
  end

  drive = nil
  if File.exists? CACHED_API_FILE
    File.open(CACHED_API_FILE) do |file|
      drive = Marshal.load(file)
    end
  else
    drive = client.discovered_api('drive', API_VERSION)
    File.open(CACHED_API_FILE, 'w') do |file|
      Marshal.dump(drive, file)
    end
  end

  return client, drive
end

def get_files(client, drive)
  result = client.execute(
    api_method: drive.files.list,
    parameters: {
      maxResults: 1000,
    },
  )
  # jj result.data.to_hash
  result
end

class OutputAdapter
  attr_accessor :data

  def initialize
    @data = []
  end

  # object[row, col] = cel_value
  def []=(row, col, val)
    @data[row.to_i]           = [] if @data[row.to_i].blank?
    @data[row.to_i][col.to_i] = val
  end

  # abstract
  def save

  end
end

class ExcelOutputAdapter < OutputAdapter
  attr_accessor :file_name

  def initialize(file_name)
    super()
    @file_name = file_name
  end

  # Excelの場合は改行が入る
  def []=(row, col, val)
    @data[row.to_i] = [] if @data[row.to_i].blank?
    if val.instance_of? Array
      @data[row.to_i][col.to_i] = val.join("\n")
    else
      @data[row.to_i][col.to_i] = val.to_s
    end
  end

  def save
    Spreadsheet.client_encoding = 'UTF-8'
    book                        = Spreadsheet::Workbook.new
    sheet                       = book.create_worksheet
    sheet.row(0).concat %w{title kind mimeType id owner permissions}
    row_index = 1
    puts data.inspect
    @data.each do |row|
      sheet[row_index, 0] = row[0]
      sheet[row_index, 1] = row[1]
      sheet[row_index, 2] = row[2]
      sheet[row_index, 3] = row[3]
      sheet[row_index, 4] = row[4]
      sheet[row_index, 5] = row[5]
      row_index           += 1
    end
    book.write @file_name
  end
end

class TsvOutputAdapter < OutputAdapter
  attr_accessor :file_name

  def initialize(file_name)
    super()
    @file_name = file_name
  end

  # TSVでは改行が入らないのでカンマ区切りで突っ込む
  def []=(row, col, val)
    @data[row.to_i] = [] if @data[row.to_i].blank?
    if val.instance_of? Array
      @data[row.to_i][col.to_i] = ''
      val.each do |v|
        @data[row.to_i][col.to_i] << v.gsub(/[\r\n]/, '')
        @data[row.to_i][col.to_i] << ","
      end

    else
      @data[row.to_i][col.to_i] = val.to_s
    end
  end

  def save
    File.open(@file_name, 'w') do |fp|
      fp.write (%w{title kind mimeType id owner permissions}).join("\t") + "\n"
      @data.each do |row|
        fp.write row[0] + "\t"
        fp.write row[1] + "\t"
        fp.write row[2] + "\t"
        fp.write row[3] + "\t"
        fp.write row[4] + "\t"
        fp.write row[5] + "\n"
      end
    end
  end
end

# begin parse options
opt = OptionParser.new

opt.on('-v', '--verbose') do |v|
  debug = true
end

file_name = nil
opt.on('-f FILENAME') do |name|
  file_name = name
end

output = nil
opt.on('--type [TYPE]') do |type|
  case type
  when 'excel'
    output = ExcelOutputAdapter.new(file_name.presence || 'result.xls')
  when 'tsv'
    output = TsvOutputAdapter.new(file_name.presence || 'result.tsv')
  else
    puts 'You must specify file type options ""--type (xls|tsv) "'
    exit
  end
end

filter_string = nil
opt.on('--only-includes NAME') do |name|
  filter_string = name
end

opt.parse!(ARGV)
# end parse options

client, drive = setup()

row_index        = 0
all_files_result = get_files(client, drive)
all_files_result.data.items.each do |file|
  if debug
    STDERR.puts "fetching id: #{file.id}, title: #{file.title}..."
  end

  # get owners
  owners = []
  file.owners.each do |owner|
    owners << "#{owner.try(:display_name)} <#{owner.try(:email_address)}>"
  end

  # get permissions
  permission_result = client.execute(
    :api_method => drive.permissions.list,
    :parameters => { 'fileId' => file.id }
  )
  permissions       = []
  permission_result.data.items.each do |permission|
    permissions << "#{permission.role}:#{permission.name} <#{permission.try(:email_address)}>"
  end

  if filter_string.present?
    available = false
    [owners, permissions].flatten.each do |str|
      if str.include?(filter_string)
        available = true
      end
    end
    next unless available
  end

  output[row_index, 0] = file.title
  output[row_index, 1] = file.kind
  output[row_index, 2] = file.mimeType
  output[row_index, 3] = file.id
  output[row_index, 4] = owners
  output[row_index, 5] = permissions
  row_index            += 1
end

output.save

