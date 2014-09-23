# XlsxImporter

I really like [Smarter CSV](https://github.com/tilo/smarter_csv) but as MS Excel cannot export CSV files in unicode,
I need to import the spreadsheet directly to preserve the unicode characters.

XlsxImporter is a port from Smarter CSV with a few specific options for spreadsheets import. It uses [SimpleXlsxReader](https://github.com/woahdae/simple_xlsx_reader)
to read the MS Excel xlsx file. It only support xlsx files.

## Installation

Add this line to your application's Gemfile:

    gem 'xlsx_importer'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install xlsx_importer

## Usage

```ruby
    > XlsxImporter.process('/tmp/people.xlsx', options = {})
     => [ {:first_name=>"Dan", :last_name=>"McAllister"},
          {:first_name=>"Lucy", :last_name=>"Laweless"} ]
```

### Options

A few options from Smarter CVS are supported, a few more were added.

| Option                    | Default | Explanation
|---------------------------|---------|-----------------------------------------------------------------------------
| :key_mapping_hash         | nil     | a hash which maps headers from the spreadsheet to keys in the result hash
| :remove_unmapped_keys     | false   | when using :key_mapping option, should non-mapped keys / columns be removed?
| :sheet                    | 0       | the number of the sheet inside the workbook to parse
| :strip_whitespace         | true    | remove whitespace before/after values and headers
| :downcase_header          | true    | downcase all column headers
| :strip_chars_from_headers | nil     | RegExp to remove extraneous characters from the header line
| :remove_empty_hashes      | true    | remove / ignore any hashes which don't have any key/value pairs
| :remove_empty_values      | true    | remove values which have nil or empty strings as values
| :remove_zero_values       | false   | remove values which have a numeric value equal to zero / 0
| :remove_values_matching   | nil     | removes key/value pairs if value matches given regular expressions
| :header_row               | 1       | row number with the headers
| :date_keys                | nil     | Array with the keys of date fields to perform DB compatible validation

### Example

```ruby
  result = XlsxImporter.process('tmp/test.xlsx', {
      :key_mapping_hash => {
          :name => :full_name,
          :date_of_creation => :created_at,
          :date_of_termination => :deleted_at,
      },
      :date_keys => [:created_at, :deleted_at],
      :remove_unmapped_keys => true,
      :strip_chars_from_headers => /[\/\(\)]/,
      :remove_values_matching => /^[\.\-]$/,
      :sheet => 1,
      :header_row => 3
  })
```

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
