require "./spec_helper"

Spectator.describe Xls do
  include Helpers

  describe Xls::Spreadsheet do
    describe "self.xls_version" do
      it "matches version pattern" do
        expect(described_class.xls_version).to match(/\d\.\d\.\d/)
      end
    end

    describe "self.open_file" do
      it "initializes using LibXls.open_file" do
        expect(described_class.open_file(test_fixture)).to be_a(described_class)
      end
    end

    describe "self.new" do
      context "with file path" do
        it "initializes" do
          expect(described_class.new(test_fixture)).to be_a(described_class)
        end
      end

      context "with file content as string" do
        it "initializes" do
          content = File.read(test_fixture)
          expect(described_class.new(content)).to be_a(described_class)
        end
      end

      context "with file content as io" do
        it "initializes" do
          File.open(test_fixture) do |file|
            expect(described_class.new(file)).to be_a(described_class)
          end
        end
      end
    end

    describe "self.open" do
      it "yields a Workbook" do
        described_class.open(test_fixture) do |wb|
          expect(wb).to be_a(Xls::Workbook)
        end
      end
    end

    describe "#close!" do
      double Xls::Spreadsheet do
        stub close! : Nil
      end

      it "closes the workbook" do
        dbl = double(Xls::Spreadsheet)
        allow(dbl).to receive(:close!).and_return(nil)
        expect(dbl.close!).to be_nil
      end
    end

    describe "#workbook!" do
      alias XlsWorkBook = Xls::LibXls::XlsWorkBook

      it "returns a raw pointer" do
        instance = described_class.new(test_fixture)
        expect(instance.workbook!).to be_a(Pointer(XlsWorkBook))
      end
    end
  end

  describe Xls::Workbook do
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
end
