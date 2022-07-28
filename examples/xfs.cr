require "../src/xls"

# Example: list xls extended formats (XF) per worksheet
#
# Usage: [program] <xls file path>
Xls::Spreadsheet.open(Path.new(ARGV[0])) do |s|
  s.worksheets.each do |ws|
    s.xfs.each do |xf|
      puts xf
      puts
    end
  end
end
