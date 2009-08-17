require File.join(File.expand_path(File.dirname(__FILE__)), "..", "spec_helper")

describe Poiyer::Excel::Workbook do
  before :each do
    @wb = Poiyer::Excel::Workbook.new
  end

  it "should create a workbook" do
    @wb.class.should equal(Poiyer::Excel::Workbook)
  end

  describe "creating a style" do
    it "should handle a background color" do
      color = :lime
      style = @wb.create_style(:bg_color => color)
      style.fill_foreground_color == Poiyer::Excel::Color::LIME.index
    end

    it "should handle a builtin data format" do
      format = "($#,##0_);[Red]($#,##0)"
      style = @wb.create_style(:format => format)
      style.data_format_string.should == format
    end

    it "should handle named builtin format" do
      format = :currency_precision_colored
      style = @wb.create_style(:format => format)
      style.data_format_string.should == Poiyer::Excel::DataFormat::NAMED_BUILTINS[format]
    end

    describe "for fonts" do
      it "should handle color" do
        color = :blue
        style = @wb.create_style(:color => color)
        style.get_font(@wb).color.should == Poiyer::Excel::Color.index_by_name(color)
      end

      it "should handle bold by boolean" do
        style = @wb.create_style(:bold => true)
        style.get_font(@wb).boldweight.should == Poiyer::Excel::Font::BOLDWEIGHT_BOLD
      end

      it "should handle bold by weight (integer)" do
        bold = 100
        style = @wb.create_style(:bold => bold)
        style.get_font(@wb).boldweight.should == bold
      end

      it "should handle italic" do
        style = @wb.create_style(:italic => true)
        style.get_font(@wb).italic.should == true
      end

      it "should handle strikeout" do
        style = @wb.create_style(:strikeout => true)
        style.get_font(@wb).strikeout.should == true
      end

      it "should handle underline" do
        style = @wb.create_style(:underline => true)
        style.get_font(@wb).underline.should == Poiyer::Excel::Font::U_SINGLE
      end
    end

    it "should not effect other cells" do
      style = @wb.create_style(:bold => true)
      data = "text"
      ws = @wb.create_sheet
      ws.write(0, 0, data, style)
      ws.write(0, 1, data)
      unstyled_cell_style = ws[0].cell(1).cell_style
      unstyled_cell_style.get_font(@wb).boldweight.should == Poiyer::Excel::Font::BOLDWEIGHT_NORMAL
    end

  end

end

describe Poiyer::Excel::Workbook do
  include Poiyer::Helper

  it "should save to a specified filename" do
    cleanup_test_file
    wb = Poiyer::Excel::Workbook.new
    wb.save(filename)
    File.exists?(filename).should == true
    cleanup_test_file
  end

end

describe Poiyer::Excel::Worksheet, "given a workbook" do
  before :each do
    @wb = Poiyer::Excel::Workbook.new
  end

  it "should create a sheet" do
    ws = @wb.create_sheet
    ws.class.should equal(Poiyer::Excel::Worksheet)
  end

  it "should write nil to a cell" do
    ws = @wb.create_sheet
    data = nil
    ws[1][2] = data
    ws[1][2].should == data
  end

  it "should write a boolean to a cell" do
    ws = @wb.create_sheet
    data = false
    ws[1][2] = data
    ws[1][2].should == data
  end

  it "should write a string to a cell" do
    ws = @wb.create_sheet
    data = "elephant"
    ws.write(0, 0, data)
    ws[0][0].should == data
  end

  it "should write a number to a cell" do
    ws = @wb.create_sheet
    data = 1.0
    ws.write(0, 0, data)
    ws[0][0].should == data
  end

  it "should handle an array give to write" do
    ws = @wb.create_sheet
    data = [1.0, 2.0]
    ws.write(0, 0, data)
    ws[0][0].should == data.first
    ws[0][1].should == data.last
  end

  it "should write data to a row" do
    ws = @wb.create_sheet
    data = ["red", "orange", "yellow", "green", "blue", "indigo", "violet"]
    ws.write_row(0, 0, data)
    data.each_with_index do |color, index|
      ws[0][index].should == color
    end
  end

  it "should write data to a column" do
    ws = @wb.create_sheet
    data = ["gold", "silver", "bronze"]
    ws.write_column(0, 0, data)
    data.each_with_index do |metal, index|
      ws[index][0].should == metal
    end
  end

  it "should merge a region" do
    ws = @wb.create_sheet
    row = [0,1,2,3]
    column = [0,1,2,3]
    region_num = ws.merge!(row.first, row.last, column.first, column.last)
    ws.get_num_merged_regions.should == 1
    region = ws.get_merged_region_at(region_num)
    region.row_from.should == row.first
    region.row_to.should == row.last
    region.column_from.should == column.first
    region.column_to.should == column.last
  end

  it "should freeze a row" do
    ws = @wb.create_sheet
    ws.freeze_row!
  end

  it "should freeze a column" do
    ws = @wb.create_sheet
    ws.freeze_column!
  end

  it "should freeze a row and column" do
    ws = @wb.create_sheet
    ws.freeze!(2, 4)
  end

  it "should auto-size all columns on a sheet" do
    ws = @wb.create_sheet
    column = 2
    old_column_width = ws.get_column_width(column)
    ws.write(3, column, "AFSDFSFLSKFJSLDKFJSLDFJSLDFJSLFJSLFJ")
    ws.auto_size!
    ws.get_column_width(column).should > old_column_width
  end
end

describe "Support for 'gem spreadsheet'" do
  describe Poiyer::Excel::Workbook do
    include Poiyer::Helper

    it "should save to a specified file given a filename with new" do
      cleanup_test_file
      wb = Poiyer::Excel::Workbook.new(filename)
      wb.close
      File.exists?(filename).should == true
      cleanup_test_file
    end
  end

  describe Poiyer::Excel::Worksheet do
    it "should add a named worksheet" do
      sheet_name = "Tab Title"
      wb = Poiyer::Excel::Workbook.new
      ws = wb.add_worksheet(sheet_name)
      ws.class.should equal(Poiyer::Excel::Worksheet)
      ws.should equal(wb.get_sheet(sheet_name))
    end
  end
end
