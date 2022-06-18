require "../src/xls"
require "spectator"

module Helpers
  def fixture_path(filename)
    Path.new("#{Dir.current}/spec/fixtures/#{filename}")
  end

  def test_fixture
    fixture_path("test.xls")
  end
end
