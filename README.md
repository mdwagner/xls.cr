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

## Contributors

- [Michael Wagner](https://github.com/mdwagner) - creator and maintainer
