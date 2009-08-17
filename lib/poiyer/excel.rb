module Poiyer
  module Excel
    FileOutputStream = java.io.FileOutputStream
    Workbook   = org.apache.poi.hssf.usermodel.HSSFWorkbook
    Worksheet  = org.apache.poi.hssf.usermodel.HSSFSheet
    Row        = org.apache.poi.hssf.usermodel.HSSFRow
    Region     = org.apache.poi.hssf.util.Region
    Cell       = org.apache.poi.hssf.usermodel.HSSFCell
    CellStyle  = org.apache.poi.hssf.usermodel.HSSFCellStyle
    Color      = org.apache.poi.hssf.util.HSSFColor
    DataFormat = org.apache.poi.hssf.usermodel.HSSFDataFormat
    Font       = org.apache.poi.hssf.usermodel.HSSFFont

    class Workbook
      alias :add_worksheet :create_sheet

      def initialize(filename = nil)
        super()
        if filename
          @file = FileOutputStream.new(filename)
        end
      end

      # Creates a new cell style object for the worksheet provided
      #  
      #     style = wb.create_style(:bg_color => :lime)
      #
      # Options Available are
      # * :bg_color
      # * :format
      # * :color
      # * :bold
      # * :italic
      # * :strikeout
      # * :underline
      def create_style(opts = {})
        CellStyle.create(self, opts)
      end

      # Saves the workbook to s specified file
      #  wb = Poiyer::Excel::Workbook.new
      #  wb.save("filename.xls")
      def save(filename = nil)
        write( FileOutputStream.new(filename) )
      end

      # Closes a workbook and writes it to the file specified
      # in the constructor if one was provided
      #  wb = Poiyer::Excel::Workbook.new("filename.xls")
      #  wb.close
      def close
        write(@file) if @file
      end

    end

    class Worksheet
      def [] (index)
        get_row(index) || create_row(index)
      end

      def write(irow, icol, data=nil, style=nil)
        case data
        when Array
          write_row(irow, icol, data, style)
        else
          cell = self[irow].cell(icol)
          cell.cell_value = data
          cell.cell_style = style if style
        end
      end

      def write_row(irow, icol, tokens = nil, style = nil)
        case tokens
        when Array
          tokens.each do |token|
            case token
            when Array
              write_column(irow, icol, token, style)
            else
              write(irow, icol, token, style)
            end
            icol += 1
          end
        else
          write(irow, icol, tokens, style)
        end
      end

      def write_column(irow, icol, tokens = nil, style = nil)
        case tokens
        when Array
          tokens.each do |token|
            case token
            when Array
              write_row(irow, icol, token, style)
            else
              write(irow, icol, token, style)
            end
            irow +=1
          end
        else
          write(irow, icol, tokens, style)
        end
      end

      def merge!(row_start, row_end, column_start, column_end)
        self.add_merged_region( Region.new(row_start, column_start, row_end, column_end) )
      end

      def freeze!(row = 0, column = 0)
        self.create_freeze_pane(column || 0, row || 0)
      end

      def freeze_row!(row = 1)
        freeze!(row)
      end

      def freeze_column!(column = 1)
        freeze!(nil, column)
      end

      def auto_size!
        longest_row = 1
        self.iterator.each do |row|
          row_length = row.last_cell_num
          longest_row = row_length if row_length > longest_row
        end
        (0..longest_row).each do |index|
          self.auto_size_column(index)
        end
      end

    end

    class Row

      # Overriding the cell method automatically created by JRuby
      def cell(index)
        get_cell(index) || self.create_cell(index)
      end

      def [] (index)
        cell = self.cell(index)
        case cell.cell_type
        when Cell::CELL_TYPE_BLANK
          nil
        when Cell::CELL_TYPE_NUMERIC
          cell.numeric_cell_value
        when Cell::CELL_TYPE_STRING
          cell.rich_string_cell_value.get_string
        when Cell::CELL_TYPE_BOOLEAN
          cell.boolean_cell_value
        when Cell::CELL_TYPE_FORUMULA
          cell.cell_formula
        when Cell::CELL_TYPE_ERROR
          cell.error_cell_value
        end
      end

      def []=(index, value)
        self.cell(index).cell_value = value
      end
    end

    class Cell
      def style
        cell_style
      end

      def to_s
        to_string
      end
    end

    class CellStyle
      def self.create(workbook, opts)
        style = workbook.create_cell_style
        style.build(workbook, opts)
      end

      def build(workbook, opts)
        if opts[:bg_color]
          self.fill_pattern = SOLID_FOREGROUND
          self.fill_foreground_color = Color.index_by_name(opts[:bg_color])
        end

        if opts[:format]
          if DataFormat::NAMED_BUILTINS.keys.include?(opts[:format])
            opts[:format] = DataFormat::NAMED_BUILTINS[opts[:format]]
          end
          self.data_format = DataFormat.get_builtin_format(opts[:format])
        end

        font_options = [:color, :bold, :italic, :strikeout, :underline]
        if opts.keys.any? { |opt| font_options.include?(opt) }
          font = workbook.create_font
          font.color = Color.index_by_name(opts[:color]) if opts.key?(:color)
          if opts.key?(:bold)
            case opts[:bold]
            when true; opts[:bold] = Font::BOLDWEIGHT_BOLD
            when false; opts[:bold] = Font::BOLDWEIGHT_NORMAL
            end
            font.boldweight = opts[:bold]
          end
          font.italic = opts[:italic] if opts.key?(:italic)
          font.strikeout = opts[:strikeout] if opts.key?(:strikeout)
          if opts.key?(:underline)
            font.underline = Font::U_SINGLE if opts[:underline] == true
          end
          self.font = font
        end

        self
      end
    end

    class DataFormat
      NAMED_BUILTINS = {
        :currency            => "($#,##0_);($#,##0)",
        :currency_colored    => "($#,##0_);[Red]($#,##0)",
        :currency_precision  => "($#,##0.00);($#,##0.00)",
        :currency_precision_colored =>"($#,##0.00_);[Red]($#,##0.00)",
        :percent             => "0%",
        :percent_precision   => "0.00%",
      }.freeze
    end

    class Color
      # Default Nested Classes in HSSFColor
      COLORS = [:black, :brown, :olive_green, :dark_green,
        :dark_teal, :dark_blue, :indigo, :grey_80_percent,
        :orange, :dark_yellow, :green, :teal, :blue,
        :blue_grey, :grey_50_percent, :red, :light_orange, :lime,
        :sea_green, :aqua, :light_blue, :violet, :grey_40_percent,
        :pink, :gold, :yellow, :bright_green, :turquoise,
        :dark_red, :sky_blue, :plum, :grey_25_percent, :rose,
        :light_yellow, :light_green, :light_turquoise, :pale_blue,
        :lavender, :white, :cornflower_blue, :lemon_chiffon,
        :maroon, :orchid, :coral, :royal_blue,
        :light_cornflower_blue, :tan].freeze

      def self.index_by_name(name)
        class_name = COLORS.first.to_s.upcase
        if COLORS.include?(name.to_sym)
          class_name = name.to_s.upcase
        end
        const_get(class_name).index
      end

    end

  end
end
