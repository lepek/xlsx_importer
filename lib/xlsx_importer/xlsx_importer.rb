require 'simple_xlsx_reader'

module XlsxImporter
  def self.process(input, options={})

    default_options = {
      :key_mapping_hash => nil,
      :remove_unmapped_keys => false,
      :headers_in_file => true,
      :sheet => 0,
      :strip_whitespace => true,
      :downcase_header => true,
      :strip_chars_from_headers => nil,
      :remove_empty_hashes => true,
      :remove_empty_values => true,
      :remove_zero_values => false,
      :remove_values_matching => nil,
      :header_row => 1,
      :date_keys => nil,
    }
    options = default_options.merge(options)

    headers = []
    result = []

    doc = SimpleXlsxReader.open(input)
    sheet = doc.sheets[options[:sheet]]

    headers = sheet.rows[options[:header_row] - 1]
    headers.map!{|x| x.gsub(options[:strip_chars_from_headers], '')} if options[:strip_chars_from_headers]
    headers.map!{|x| x.strip if x.respond_to?(:strip)}  if options[:strip_whitespace]
    headers.map!{|x| x.gsub(/\s+/,'_')}
    headers.map!{|x| x.downcase if x.respond_to?(:downcase) } if options[:downcase_header]
    headers.map!{|x| x.to_sym }

    key_mapping_hash = options[:key_mapping_hash]
    if !key_mapping_hash.nil? && key_mapping_hash.class == Hash && key_mapping_hash.keys.size > 0
      headers.map!{|x| key_mapping_hash.has_key?(x) ? (key_mapping_hash[x].nil? ? nil : key_mapping_hash[x].to_sym) : (options[:remove_unmapped_keys] ? nil : x)}
    end

    line = 0
    rows = sheet.rows.drop(options[:header_row])
    begin
      rows.each do |row|
        line += 1

        row.map!{|x| x.respond_to?(:strip) ? x.strip : x }  if options[:strip_whitespace]

        hash = Hash.zip(headers,row)

        hash.delete(nil); hash.delete('');
        hash.delete_if {|k,v| v.nil? || v =~ /^\s*$/}  if options[:remove_empty_values]
        hash.delete_if {|k,v| !v.nil? && v =~ /^(\d+|\d+\.\d+)$/ && v.to_f == 0} if options[:remove_zero_values]   # values are typically Strings!
        hash.delete_if {|k,v| v =~ options[:remove_values_matching]} if options[:remove_values_matching]

        options[:date_keys].map do |x|
            hash[x] = nil if hash.has_key?(x) && self.date_validation(hash[x]).nil?
        end if options[:date_keys].present? && options[:date_keys].class == Array && options[:date_keys].size > 0

        next if hash.empty? if options[:remove_empty_hashes]

        result << hash
      end
    rescue => e
      puts "Error around line #{line} \n"
      raise e
    end

    return result

  end

  def self.date_validation(date)
    date = Date.parse(date) if date.class == 'String'
    date.strftime('%Y-%m-%d') =~ /^([1-3][0-9]{3,3})-(0?[1-9]|1[0-2])-(0?[1-9]|[1-2][1-9]|3[0-1])$/
  rescue
    return nil
  end



end