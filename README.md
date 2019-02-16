# Rxl

The purpose of the RXL gem is to provide a ruby/Excel interface that provides the following features:

1. Specification using Excel key indices (A1, B5 etc)
2. Avoiding multi-level class management by utilising the ruby hash
3. Simplified handling with the aim of doing less, better - eg no setting of properties for full rows/columns in Excel files

The mechanics of the conversion between xlsx and ruby hash have been implemented using the RubyXL gem:

https://github.com/weshatheleopard/rubyXL

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'rxl'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install rxl

## Usage

With some exceptions (due mainly to the vagaries of Excel) a file can be read in and the resulting hash passed to the write method to save a duplicate of the original.

### Read from file

To read a file to hash simply pass the filepath:

```ruby
Rxl.read_file('path/to/file.xlxs')
```

The format of the excel read hash has the following skeleton:

```ruby
{
    'Sheet1' => {
      row_count: 1,
      column_count: 1,
      rows: {},
      columns: {},
      cells: {
          'A1' => {
              value: 'abc',
              format: 'text'
          }
      }
    }
}
```

Bear in mind the limitations of reading cell formats. Everything is read as a string other than:
* cells formatted as dates are converted to a DateTime object with the time portion set to midnight
* cells formatted as times are converted to a DateTime object with the date portion set to 31/12/1899 - unless the cell has a date prefix in which case this is carried in (this will be read as a date format as per below parsing rules)
* percentages are converted to numeric format (eg 100% = 1)
* to account for floats, which lose any trailing zeroes, decimal point information is retained in the `:decimals` value
* formulas are not read from cells with date and time formats

Within these limitations the cell hash's :format holds the best analysis of the original cell format but as there's no way to extract all of the format information directly from the sheet some information may need to be refurbished as required after import via Rxl.

Further to the above, these rules are applied by Rxl when parsing cells:
* strings are given text format
* DateTime objects are given date format except where the date is 31/12/1899 in which case they are given time format




### Read tables from file

To read a file where the data is in table format.

* headers and rows only - no concept of sums/totals or any other content is provided for.
* columns which do not contain a value in row 1 are ignored
* formatting is discarded, only the cell values are retained

```ruby
Rxl.read_file_as_tables('path/to/file.xlsx')
```

The format of the excel table read hash has the following skeleton:

```ruby
{
    'Sheet1' => [
      {
          header_a: value,
          header_b: value
      },
      {
          header_a: value,
          header_b: value
      },
    ]
}
```

### Read multiple files at once

Pass a hash of filepaths to read, get a hash of file contents back.

```ruby
filepaths_hash = {
  first_file: 'path/to/file.xlsx',
  second_file: 'path/to/file.xlsx'
}

Rxl.read_files(filepaths_hash)
Rxl.read_files(filepaths_hash, :as_tables)
```

Returns the files with sheet contents populated with hash of cells (or array of rows if :as_tables read type is specified):

```ruby
{
  first_file: {
    'Sheet1' => 'sheet_contents',
    'Sheet2' => 'sheet_contents'
  },
  second_file: {
    'Sheet1' => 'sheet_contents',
    'Sheet2' => 'sheet_contents'
  },
}
```

### Write to file

To write a file pass the filename and hash:

```ruby
Rxl.write_file('path/to/save.xlsx', hash_workbook)
```

The format of the excel hash_workbook has sheet names as keys and hashes of cells as values:

```ruby
{
    'Sheet1' => {
      'A1' => {
          value: 'abc'
      }
    }
}
```

#### Cell specification

All cells are written with the format set to general except those with a number format specified

Specify the number format according to https://support.office.com/en-us/article/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=en-US&rs=en-US&ad=US

Examples:

| value        | number format | resulting cell format | resulting cell value |
|--------------|---------------|-----------------------|----------------------|
| 0            | 0             | number                | 0                    |
| 0.49         | 0             | number                | 0                    |
| 0.5          | 0             | number                | 1                    |
| 0            | 0.00          | number                | 0.00                 |
| 0            | 0%            | percentage            | 0%                   |
| 1            | 0%            | percentage            | 100%                 |
| '01/01/2000' | 'dd/mm/yyyy'  | date                  | 01/01/2000           |

#### Write Validation

The following rules are validated for write_file:

* The hash_workbook must be a hash (NB if empty a blank file will be created with a single sheet called "Sheet1")
* The hash_workbook keys must be strings
* The hash_workbook values (hash_worksheet) must be hashes

* The hash_worksheet keys must be strings of valid Excel cell id format
* The hash_worksheet values must be hashes (specifying cells)
* The hash_cell keys must conform the the cell specification as below

* If a formula is provided the value must be nil or an empty string
* If a number format is provided the value must be consistent with it


Cells are specified as hashes following this example:

```ruby
{
  value: 'abc'
}
```

Other keys can be specified:
* v_align: :top, :centre or :bottom (default)
* h_align: :left (default), :centre or :right

TODO: add full description for hash_cell_to_rubyxl_cell and rubyxl_cell_to_hash_cell (and check they're as consistent as possible)

### Write to file as tables

To write a file as pass the filename and hash:

```ruby
Rxl.write_file_as_tables('path/to/save.xlsx', hash_tables, order)
```

The worksheets' top row will be populated with values specified in the `order` array. Those array items will also be used to extract the current row from the current hash.

* use `nil` in the `order` array to leave blank columns (including blank header)
* string or symbol keys can be used, so long as the key in order is the same as in the hashes

The format of the excel hash_workbook has sheet names as keys and hashes of rows as values:

```ruby
order = %i[header_a header_b]
hash_tables = {
  'Sheet1' => [
    {
      header_a: 'some_value',
      header_b: 'other_value'
    },
    {
      header_a: 'some_value',
      header_b: 'other_value'
    }
  ]
}
```

#### Formatting for tables

Add formatting to tables by adding a `:formats` key to the top level hash.

Inside the formatting hash add child hashes with keys for the relevant table.

Within the table add hashes for each column with the key as the column letter and the value as a cell hash (excluding `:value`).

Formatting for rows is not currently implemented.

Additionally inside the table hash add a key `:headers` with a cell hash (excluding `:value`) to set formatting for the header row.

```ruby
order = %i[header_a header_b]
hash_tables = {
  formats: {
    'Sheet1' => {
      headers: {
        bold: true,
        align: 'center'
      },
      'B' => {
        fill: 'feb302'
      }
    }
  },
  'Sheet1' => [
    {
      'col_1' => 'some_value',
      'col_2' => 'other_value'
    },
    {
      'col_1' => 'some_value',
      'col_2' => 'other_value'
    }
  ]
}
```

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/rxl. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [Contributor Covenant](http://contributor-covenant.org) code of conduct.


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).

