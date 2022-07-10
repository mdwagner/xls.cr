# xls

Crystal bindings to libxls

## Installation

1. Add the dependency to your `shard.yml`:

   ```yaml
   dependencies:
     xls:
       github: mdwagner/xls.cr
   ```

2. Run `shards install`

## Usage

```crystal
require "xls"
```

TODO: Write usage instructions here

## Development

TODO: Write development instructions here

## How it works

```
Xls::Spreadsheet(1) -> Xls::Workbook(1) -> Xls::Sheets(N) -> Xls::Worksheet(N)
```

### Xls::Spreadsheet

- Represents the `.xls` file (pointer)

### Xls::Workbook

- Metadata, including accessing sheets

### Xls::Sheets

- Parses sheets and gives you valid worksheets
- Gives you sheet names and count

### Xls::Worksheet

- Content of a worksheet (rows, cols, types, etc.)
- Types
  - work like JSON::Any/XML::Any/etc.
  - Single type is a union of all possible types
  - has different methods to assert usage: as_s/as_s?

## Contributing

1. Fork it (<https://github.com/mdwagner/xls.cr/fork>)
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request

## Contributors

- [Michael Wagner](https://github.com/mdwagner) - creator and maintainer
