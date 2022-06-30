require "../spec_helper"

Spectator.describe Xls::Workbook do
  include Helpers

  let(:yield_workbook) do
    Xls::Spreadsheet.open(test_fixture) do |wb|
      yield wb
    end
  end

  describe "#charset" do
    it "gets charset" do
      yield_workbook do |wb|
        expect(wb.charset).to eq("UTF-8")
      end
    end
  end

  describe "#summary" do
    it "gets summary" do
      yield_workbook do |wb|
        expect(wb.summary).not_to be_empty
      end
    end
  end

  describe "#doc_summary" do
    it "gets doc_summary" do
      yield_workbook do |wb|
        expect(wb.doc_summary).not_to be_empty
      end
    end
  end

  describe "#sheets" do
    it "collects sheets" do
      yield_workbook do |wb|
        expect(wb.sheets).to be_a(Xls::Sheets)
      end
    end
  end
end
