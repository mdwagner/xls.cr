# xls

Crystal bindings to [libxls](https://github.com/libxls/libxls)

## Installation

1. Add the dependency to your `shard.yml`:

   ```yaml
   dependencies:
     xls:
       github: mdwagner/xls.cr
       version: 0.2.0
   ```

2. Run `shards install`

## Usage

```crystal
require "xls"

Xls::Spreadsheet # <- library entry point
```

Look at `examples/` folder for idea of usage.

TODO: Write usage instructions here

## Development

TODO: Write development instructions here

## Contributing

1. Fork it (<https://github.com/mdwagner/xls.cr/fork>)
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request

## Contributors

- [Michael Wagner](https://github.com/mdwagner) - creator and maintainer
