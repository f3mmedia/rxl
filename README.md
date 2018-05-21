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
    "Sheet1" => {
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
* numbers (including floats and percentages) where the cell format is number or percentage are read in as integers - trailing zeroes are cropped from floats in this case and percentages are converted to numeric format (eg 100% = 1)
* formulas are not read from cells with date and time formats

Within these limitations the cell hash's :format holds the best analysis of the original cell format but as there's no way to extract all of the format information directly from the sheet some information may need to be refurbished as required after import via Rxl.

Further to the above, these rules are applied by Rxl when parsing cells:
* strings are given text format
* DateTime objects are given date format except where the date is 31/12/1899 in which case they are given time format




### Read tables from file

To read a file where the data is in table format - headers and values, no totals or otherwise extra content:

```ruby
Rxl.read_file_as_tables('path/to/file.xlsx')
```

The format of the excel table read hash has the following skeleton:

```ruby
{
    "Sheet1" => [
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

### Write to file

To write a file pass the filename and hash:

```ruby
Rxl.write_file('path/to/save.xlsx', hash_workbook)
```

The format of the excel hash_workbook must contain at least the following skeleton:

```ruby
{
    "Sheet1" => {
      cells: {
          'A1' => {
              value: 'abc',
              format: 'text'
          }
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


* The hash_worksheet keys must be symbols
* The following keys are allowed for hash_worksheet: :cells, :rows, :columns
* The hash_worksheet values must be arrays - those arrays must contain only hashes


* The arrays' child hash keys must be strings of valid Excel cell id format (or stringified number for a row, capitalised alpha for a column)
* The arrays' child hash values must be hashes (hash_cell)
* The hash_cell keys must conform the the cell specification as below

* If a formula is provided the value must be nil or an empty string
* If a number format is provided the value must be consistent with it


TODO: Add further detail

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/rxl. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [Contributor Covenant](http://contributor-covenant.org) code of conduct.


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).

