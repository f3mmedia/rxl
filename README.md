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

### Write to file

To write a file pass the filename and hash:

```ruby
Rxl.write_file('path/to/save.xlsx', write_hash)
```

The format of the excel write hash must contain at least the following skeleton:

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


TODO: Add further detail

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/rxl. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [Contributor Covenant](http://contributor-covenant.org) code of conduct.


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).

