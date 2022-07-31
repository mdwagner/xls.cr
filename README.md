# xls

Crystal bindings to [libxls](https://github.com/libxls/libxls) to read old Excel files (.xls)

## Installation

1. Install **[libxls](https://github.com/libxls/libxls)**

1. Add the dependency to your `shard.yml`:

   ```yaml
   dependencies:
     xls:
       github: mdwagner/xls.cr
       version: 0.4.1
   ```

1. Run `shards install`

## Usage

```crystal
require "xls"

# Simple example of usage
Xls::Spreadsheet.open(Path.new("./example.xls")) do |s|
  skip_first_ws_puts = true

  s.worksheets.each do |ws|
    puts unless skip_first_ws_puts
    puts "Sheet: #{ws.name}"
    skip_first_row_puts = true

    ws.rows.each_with_index do |row, row_index|
      puts unless skip_first_row_puts
      puts "Row: #{row_index + 1}"

      row.cells.each_with_index do |cell, cell_index|
        puts "Cell: #{cell_index + 1}, Value: #{cell.value}"
      end

      skip_first_row_puts = false
    end

    skip_first_ws_puts = false
  end
end
```

Look at `examples/` folder for more examples of different usage.

## Contributors

- [Michael Wagner](https://github.com/mdwagner) - creator and maintainer
