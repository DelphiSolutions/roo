require 'date'
require 'nokogiri'

class Roo::Excelx < Roo::Base
  module Format
    EXCEPTIONAL_FORMATS = {
      'h:mm am/pm' => :date,
      'h:mm:ss am/pm' => :date,
    }

    STANDARD_FORMATS = {
      0 => 'General',
      1 => '0',
      2 => '0.00',
      3 => '#,##0',
      4 => '#,##0.00',
      9 => '0%',
      10 => '0.00%',
      11 => '0.00E+00',
      12 => '# ?/?',
      13 => '# ??/??',
      14 => 'mm-dd-yy',
      15 => 'd-mmm-yy',
      16 => 'd-mmm',
      17 => 'mmm-yy',
      18 => 'h:mm AM/PM',
      19 => 'h:mm:ss AM/PM',
      20 => 'h:mm',
      21 => 'h:mm:ss',
      22 => 'm/d/yy h:mm',
      37 => '#,##0 ;(#,##0)',
      38 => '#,##0 ;[Red](#,##0)',
      39 => '#,##0.00;(#,##0.00)',
      40 => '#,##0.00;[Red](#,##0.00)',
      45 => 'mm:ss',
      46 => '[h]:mm:ss',
      47 => 'mmss.0',
      48 => '##0.0E+0',
      49 => '@',
    }

    def to_type(format)
      format = format.to_s.downcase
      if type = EXCEPTIONAL_FORMATS[format]
        type
      elsif format.include?('#')
        :float
      elsif format.include?('d') || format.include?('y')
        if format.include?('h') || format.include?('s')
          :datetime
        else
          :date
        end
      elsif format.include?('h') || format.include?('s')
        :time
      elsif format.include?('%')
        :percentage
      else
        :float
      end
    end

    module_function :to_type
  end

  ExceedsMaxError = Class.new(StandardError)

  # initialization and opening of a spreadsheet file
  # values for packed: :zip
  def initialize(filename, options = {}, deprecated_file_warning = :error)
    if Hash === options
      packed = options[:packed]
      file_warning = options[:file_warning] || :error
      cell_max = options[:cell_max]
      minimal_load = options[:minimal_load] || false
    else
      warn 'Supplying `packed` or `file_warning` as separate arguments to `Roo::Excelx.new` is deprecated. Use an options hash instead.'
      packed = options
      file_warning = deprecated_file_warning
      minimal_load = false
    end

    file_type_check(filename,'.xlsx','an Excel-xlsx', file_warning, packed)
    make_tmpdir do |tmpdir|
      filename = download_uri(filename, tmpdir) if uri?(filename)
      filename = unzip(filename, tmpdir) if packed == :zip
      @filename = filename
      unless File.file?(@filename)
        raise IOError, "file #{@filename} does not exist"
      end
      @comments_files = Array.new
      extract_content(tmpdir, @filename)
      # We have to encode to UTF-8 because nokogiri explodes gsubing UTF-16/other encoding headers (bug?)
      @workbook_doc = load_xml(File.join(tmpdir, "roo_workbook.xml"), "UTF-8")

      if cell_max
        cell_count = Roo::Base.cells_in_range(dimensions)
        raise ExceedsMaxError, "Excel file exceeds cell maximum: #{cell_count} > #{cell_max}" if cell_count > cell_max
      end

      @shared_table = []
      if File.exist?(File.join(tmpdir, 'roo_sharedStrings.xml'))
        @sharedstring_doc = load_xml(File.join(tmpdir, 'roo_sharedStrings.xml'), "UTF-8")
        read_shared_strings(@sharedstring_doc)
      end
      @styles_table = []
      @style_definitions = Array.new # TODO: ??? { |h,k| h[k] = {} }
      if File.exist?(File.join(tmpdir, 'roo_styles.xml'))
        @styles_doc = load_xml(File.join(tmpdir, 'roo_styles.xml'), "UTF-8")
        read_styles(@styles_doc)
      end
      if minimal_load == true
        #warn ':minimal_load option will not load ANY sheets, or comments into memory, do not use unless you are
        #      ONLY interested in using each_row to iterate over a sheet and extract values for each row'
      else
        @sheet_doc = @sheet_files.compact.map do |item|
          load_xml(item, "UTF-8")
        end
        @comments_doc = @comments_files.compact.map do |item|
          load_xml(item, "UTF-8")
        end
      end
    end
    super(filename, options)
    @formula = Hash.new
    @excelx_type = Hash.new
    @excelx_value = Hash.new
    @s_attribute = Hash.new # TODO: ggf. wieder entfernen nur lokal benoetigt
    @comment = Hash.new
    @comments_read = Hash.new
  end

  def method_missing(m,*args)
    # is method name a label name
    read_labels
    if @label.has_key?(m.to_s)
      sheet ||= @default_sheet
      read_cells(sheet)
      row,col = label(m.to_s)
      cell(row,col)
    else
      # call super for methods like #a1
      super
    end
  end

  # Returns the content of a spreadsheet-cell.
  # (1,1) is the upper left corner.
  # (1,1), (1,'A'), ('A',1), ('a',1) all refers to the
  # cell at the first line and first row.
  def cell(row, col, sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    row, col = normalize(row, col)

    @cell[sheet][[row, col]]
  end

  # Returns the formula at (row,col).
  # Returns nil if there is no formula.
  # The method #formula? checks if there is a formula.
  def formula(row,col,sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    @formula[sheet][[row,col]] && @formula[sheet][[row,col]]
  end
  alias_method :formula?, :formula

  # returns each formula in the selected sheet as an array of elements
  # [row, col, formula]
  def formulas(sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    if @formula[sheet]
      @formula[sheet].each.collect do |elem|
        [elem[0][0], elem[0][1], elem[1]]
      end
    else
      []
    end
  end

  class Font
    attr_accessor :bold, :italic, :underline

    def bold?
      @bold == true
    end

    def italic?
      @italic == true
    end

    def underline?
      @underline == true
    end
  end

  # Given a cell, return the cell's style
  def font(row, col, sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    s_attribute = @s_attribute[sheet][[row,col]]
    s_attribute ||= 0
    s_attribute = s_attribute.to_i
    @style_definitions[s_attribute]
  end

  # returns the type of a cell:
  # * :float
  # * :string,
  # * :date
  # * :percentage
  # * :formula
  # * :time
  # * :datetime
  def celltype(row,col,sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    if @formula[sheet][[row,col]]
      return :formula
    else
      @cell_type[sheet][[row,col]]
    end
  end

  # returns the internal type of an excel cell
  # * :numeric_or_formula
  # * :string
  # Note: this is only available within the Excelx class
  def excelx_type(row,col,sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    return @excelx_type[sheet][[row,col]]
  end

  # returns the internal value of an excelx cell
  # Note: this is only available within the Excelx class
  def excelx_value(row,col,sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    return @excelx_value[sheet][[row,col]]
  end

  # returns the internal format of an excel cell
  def excelx_format(row,col,sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    s = @s_attribute[sheet][[row,col]]
    attribute2format(s).to_s
  end

  # returns an array of sheet names in the spreadsheet
  def sheets
    @sheet_names ||= @workbook_doc.xpath("//xmlns:sheet").map do |sheet|
      sheet['name']
    end
  end

  # returns the sheet index by name
  # passed indexes through
  def sheet_index(sheet)
    if sheet.is_a?(Fixnum)
      sheet
    else
      sheets.index(sheet)
    end
  end

  # shows the internal representation of all cells
  # for debugging purposes if debug is truthy
  def to_s(sheet = nil, debug = nil)
    sheet ||= @default_sheet
    # only print the whole thing on debug
    if debug
      read_cells(sheet)
      @cell[sheet].inspect
    else
      super()
    end
  end

  def inspect(sheet = nil, debug = nil)
    to_s(sheet, debug)
  end

  # returns the row,col values of the labelled cell
  # (nil,nil) if label is not defined
  def label(labelname)
    read_labels
    if @label.empty? || !@label.has_key?(labelname)
      return nil,nil,nil
    else
      return @label[labelname][1].to_i,
        Roo::Base.letter_to_number(@label[labelname][2]),
        @label[labelname][0]
    end
  end

  # Returns an array which all labels. Each element is an array with
  # [labelname, [row,col,sheetname]]
  def labels
    # sheet ||= @default_sheet
    # read_cells(sheet)
    read_labels
    @label.map do |label|
      [ label[0], # name
        [ label[1][1].to_i, # row
          Roo::Base.letter_to_number(label[1][2]), # column
          label[1][0], # sheet
        ] ]
    end
  end

  # returns the comment at (row/col)
  # nil if there is no comment
  def comment(row,col,sheet=nil)
    sheet ||= @default_sheet
    #read_cells(sheet)
    read_comments(sheet) unless @comments_read[sheet]
    row,col = normalize(row,col)
    return nil unless @comment[sheet]
    @comment[sheet][[row,col]]
  end

  # true, if there is a comment
  def comment?(row,col,sheet=nil)
    sheet ||= @default_sheet
    # read_cells(sheet)
    read_comments(sheet) unless @comments_read[sheet]
    row,col = normalize(row,col)
    comment(row,col) != nil
  end

  # returns each comment in the selected sheet as an array of elements
  # [row, col, comment]
  def comments(sheet=nil)
    sheet ||= @default_sheet
    read_comments(sheet) unless @comments_read[sheet]
    if @comment[sheet]
      @comment[sheet].each.collect do |elem|
        [elem[0][0],elem[0][1],elem[1]]
      end
    else
      []
    end
  end

  # Get the dimensions of a given sheet without parsing the whole
  # sheet. This is useful for determining if a file is realistically
  # consumable without having to parse a large xml file.
  def dimensions(options = {})
    sheet = options[:sheet] || @default_sheet
    sheet_idx = sheet_index(sheet)
    row_count = 0
    make_tmpdir do |tmpdir|
      file = extract_sheet_at_index(tmpdir, sheet_idx)
      Nokogiri::XML::Reader(File.open(file), nil, 'UTF-8').each do |node|
        next if node.name      != 'dimension'
        next if node.node_type != Nokogiri::XML::Reader::TYPE_ELEMENT
        dimension = node.attribute('ref')
        return dimension
      end
    end
  end

  # This method iterates over all the rows, loading
  # then from either read rows or streaming them
  # depending on the :minimal_load initialization
  # option.
  #
  # options[:sheet] can be used to specify a sheet other
  # than the default
  # options[:max_rows] break when parsed max rows
  # options[:cells] indicates if cells or values should be
  # returned (defaults false)
  # options[:strip_nils] removes trailing nils from rows
  def each_row(options = {})
    return enum_for(:each_row, options) unless block_given?

    row_count = 0
    if @sheet_doc
      each_row_loaded(options)
    else
      each_row_streaming(options)
    end.each do |row|
      # Handle cell and value row elements when popping
      row.pop while options[:strip_nils] && (row.last.value.nil? rescue row.last.nil?) && !row.empty?
      yield row
      row_count += 1
      break if options[:max_rows] && row_count == options[:max_rows]
    end
  end

  private

  # Used in each_row when :minimal_load option is set
  def each_row_streaming(options)
    return enum_for(:each_row_streaming, options) unless block_given?

    raise StandardError, "Documents already loaded, turn on minimal_load to use this method" if @sheet_doc

    sheet = options[:sheet] || @default_sheet
    last_row_pos = 0
    make_tmpdir do |tmpdir|
      file = extract_sheet_at_index(tmpdir, sheet_index(sheet))
      raise "Invalid sheet" unless File.exists?(file)
      Nokogiri::XML::Reader(File.open(file), nil, 'UTF-8').each do |node|
        next if node.name      != 'row'
        next if node.node_type != Nokogiri::XML::Reader::TYPE_ELEMENT

        row = Nokogiri::XML(node.outer_xml).root
        row_pos = xml_row_pos(row)
        # Add blank rows
        (last_row_pos + 1).upto(row_pos - 1).each { |_r| yield [] }
        last_row_pos = row_pos
        yield cells_for_row_element(row, options)
      end
    end
  end

  # Used in each_row when :minimal_load option is not set
  def each_row_loaded(options)
    return enum_for(:each_row_loaded, options) unless block_given?

    raise StandardError, "Documents not loaded, turn off minimal_load to use this method" unless @sheet_doc

    sheet = options[:sheet] || @default_sheet
    sheet_idx = sheet_index(sheet)
    last_row_pos = 0
    raise StandardError, "Invalid sheet" unless @sheet_doc[sheet_idx]
    @sheet_doc[sheet_idx].xpath("/xmlns:worksheet/xmlns:sheetData/xmlns:row").each do |row|
      row_pos = xml_row_pos(row)
      # Add blank rows
      (last_row_pos + 1).upto(row_pos - 1).each { |_r| yield [] }
      last_row_pos = row_pos
      yield cells_for_row_element(row, options)
    end
  end

  def xml_row_pos(row)
    row['r'].to_i
  end

  # helper function to set the internal representation of cells
  def set_cell_values(sheet,x,y,i,v,value_type,formula,
      excelx_type=nil,
      excelx_value=nil,
      s_attribute=nil)
    key = [y, x + i]
    @cell_type[sheet] ||= {}
    @cell_type[sheet][key] = value_type
    @formula[sheet] ||= {}
    @formula[sheet][key] = formula  if formula
    @cell[sheet] ||= {}
    @cell[sheet][key] = v
    @excelx_type[sheet] ||= {}
    @excelx_type[sheet][key] = excelx_type
    @excelx_value[sheet] ||= {}
    @excelx_value[sheet][key] = excelx_value
    @s_attribute[sheet] ||= {}
    @s_attribute[sheet][key] = s_attribute
  end

  def create_datetime_from(datetime_string)
    date_part, time_part = round_time_from(datetime_string).split(' ')
    yyyy, mm, dd = date_part.split('-')
    hh, mi, ss = time_part.split(':')
    DateTime.civil(yyyy.to_i, mm.to_i, dd.to_i, hh.to_i, mi.to_i, ss.to_i)
  end

  def round_time_from(datetime_string)
    date_part,time_part = datetime_string.split(' ')
    yyyy, mm, dd = date_part.split('-')
    hh, mi, ss = time_part.split(':')
    Time.new(yyyy.to_i, mm.to_i, dd.to_i, hh.to_i, mi.to_i, ss.to_r).round(0).strftime("%Y-%m-%d %H:%M:%S")
  end

  # internal cell model to help keep track of all of a cells
  # associated values when being parsed, etc.
  class Cell
    attr_accessor :coordinate, :value, :excelx_value, :excelx_type, :s_attribute, :formula, :type, :sheet

    def initialize(coordinate, value, options = {})
      @coordinate = coordinate
      @value =  value
      @excelx_value = options[:excelx_value]
      @excelx_type = options[:excelx_type]
      @s_attribute = options[:s_attribute]
      @formula = options[:formula]
      @type = options[:type]
      @sheet = options[:sheet]
      @base_data = options[:base_date]
    end

    class Coordinate
      attr_accessor :x, :y

      def initialize(x, y)
        @x = x
        @y = y
      end
    end
  end

  # read all cells in the selected sheet
  def read_cells(sheet=nil)
    sheet ||= @default_sheet
    validate_sheet!(sheet)
    sheet_idx = sheet_index(sheet)

    raise StandardError, "Documents not loaded, turn off minimal_load to use this method" unless @sheet_doc
    raise StandardError, "Invalid sheet" unless @sheet_doc[sheet_idx]

    return if @cells_read[sheet]
    @sheet_doc[sheet_index(sheet)].xpath("/xmlns:worksheet/xmlns:sheetData/xmlns:row/xmlns:c").each do |c|
      cell = parse_cell(c, sheet: sheet)
      set_cell_values(cell.sheet, cell.coordinate.x, cell.coordinate.y, 0, cell.value, cell.type,
                      cell.formula, cell.excelx_type, cell.excelx_value, cell.s_attribute) unless cell == 0
    end

    @cells_read[sheet] = true
    # begin comments
=begin
Datei xl/comments1.xml
  <?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
  <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <authors>
      <author />
    </authors>
    <commentList>
      <comment ref="B4" authorId="0">
        <text>
          <r>
            <rPr>
              <sz val="10" />
              <rFont val="Arial" />
              <family val="2" />
            </rPr>
            <t>Kommentar fuer B4</t>
          </r>
        </text>
      </comment>
      <comment ref="B5" authorId="0">
        <text>
          <r>
            <rPr>
            <sz val="10" />
            <rFont val="Arial" />
            <family val="2" />
          </rPr>
          <t>Kommentar fuer B5</t>
        </r>
      </text>
    </comment>
  </commentList>
  </comments>
=end
=begin
    if @comments_doc[sheet_index(sheet)]
      read_comments(sheet)
    end
=end
    #end comments
  end

  # fetch cells for a given row XML
  def cells_for_row_element(row_element, options = {})
    return [] unless row_element
    cell_col = 0
    row = Array.new
    row_element.children.each do |cell_element|
      cell = parse_cell(cell_element, options)
      cell_col = cell.coordinate.x
      row[cell_col - 1] = options[:cells] ? cell : cell.value
    end

    row
  end

  # parse an individual cell xml element
  def parse_cell(cell_element, options)
    s_attribute = cell_element['s'].to_i   # should be here
                                           # c: <c r="A5" s="2">
                                           # <v>22606</v>
                                           # </c>, format: , tmp_type: float
    value_type =
        case cell_element['t']
        when 's'
          :shared
        when 'b'
          :boolean
        # 2011-02-25 BEGIN
        when 'str'
          :string
        # 2011-02-25 END
        # 2011-09-15 BEGIN
        when 'inlineStr'
          :inlinestr
        # 2011-09-15 END
        else
          format = attribute2format(s_attribute)
          Format.to_type(format)
        end
    formula = nil
    y, x = Roo::Base.split_coordinate(cell_element['r'])
    cell_element.children.each do |cell|
      case cell.name
      when 'is'
        cell.children.each do |is|
          if is.name == 't'
            inlinestr_content = is.content
            value_type = :string
            v = inlinestr_content
            excelx_type = :string
            excelx_value = inlinestr_content #cell.content
            return Cell.new(Cell::Coordinate.new(x, y), v,
                            excelx_value: excelx_value,
                            excelx_type: excelx_type,
                            formula: formula,
                            type: value_type,
                            s_attribute: s_attribute,
                            sheet: options[:sheet],
                            base_date: base_date)
          end
        end
      when 'f'
        formula = cell.content
      when 'v'
        if [:time, :datetime].include?(value_type) && cell.content.to_f >= 1.0
          value_type =
              if (cell.content.to_f - cell.content.to_f.floor).abs > 0.000001
                :datetime
              else
                :date
              end
        end
        excelx_type = [:numeric_or_formula, format.to_s]
        excelx_value = cell.content
        v = case value_type
            when :shared
              value_type = :string
              excelx_type = :string
              @shared_table[cell.content.to_i]
            when :boolean
              (excelx_value.to_i == 1 ? 'TRUE' : 'FALSE')
            when :date
              yyyy, mm, dd = (base_date + excelx_value.to_i).strftime("%Y-%m-%d").split('-')
              Date.new(yyyy.to_i, mm.to_i, dd.to_i)
            when :time
              excelx_value.to_f*(24*60*60)
            when :datetime
              create_datetime_from((base_date + excelx_value.to_f.round(6)).strftime("%Y-%m-%d %H:%M:%S.%N"))
            when :formula
              excelx_value.to_f #TODO: !!!!
            when :string
              excelx_type = :string
              excelx_value
            when :percentage
              excelx_value.to_f
            else
              val = excelx_value.to_f rescue excelx_value
              val = val.to_i if (val.to_f % 1).zero? rescue val
              value_type = val.is_a?(Numeric) ? :float : :string
              val
            end
        return Cell.new(Cell::Coordinate.new(x, y), v,
                        excelx_value: excelx_value,
                        excelx_type: excelx_type,
                        formula: formula,
                        type: value_type,
                        s_attribute: s_attribute,
                        sheet: options[:sheet],
                        base_date: base_date)
      end
    end
    Cell.new(Cell::Coordinate.new(x, y), nil)
  end

  # Reads all comments from a sheet
  def read_comments(sheet=nil)
    sheet ||= @default_sheet
    validate_sheet!(sheet)
    n = sheet_index(sheet)
    return unless @comments_doc[n] #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    @comments_doc[n].xpath("//xmlns:comments/xmlns:commentList/xmlns:comment").each do |comment|
      ref = comment.attributes['ref'].to_s
      row,col = Roo::Base.split_coordinate(ref)
      comment.xpath('./xmlns:text/xmlns:r/xmlns:t').each do |text|
        @comment[sheet] ||= {}
        @comment[sheet][[row,col]] = text.text
      end
    end
    @comments_read[sheet] = true
  end

  def read_labels
    @label ||= Hash[@workbook_doc.xpath("//xmlns:definedName").map do |defined_name|
      # "Sheet1!$C$5"
      sheet, coordinates = defined_name.text.split('!$', 2)
      col,row = coordinates.split('$')
      [defined_name['name'], [sheet,row,col]]
    end]
  end

  # Extracts all needed files from the zip file
  def process_zipfile(tmpdir, zipfilename, zip, path='')
    @sheet_files = []
    Roo::ZipFile.open(zipfilename) {|zf|
      zf.entries.each {|entry|
        if entry.to_s.end_with?('workbook.xml')
          open(tmpdir+'/'+'roo_workbook.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        # if entry.to_s.end_with?('sharedStrings.xml')
        # at least one application creates this file with another (incorrect?)
        # casing. It doesn't hurt, if we ignore here the correct casing - there
        # won't be both names in the archive.
        # Changed the casing of all the following filenames.
        if entry.to_s.downcase.end_with?('sharedstrings.xml')
          open(tmpdir+'/'+'roo_sharedStrings.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        if entry.to_s.downcase.end_with?('styles.xml')
          open(tmpdir+'/'+'roo_styles.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        if entry.to_s.downcase =~ /sheet([0-9]+).xml$/
          nr = $1
          open(tmpdir+'/'+"roo_sheet#{nr}",'wb') {|f|
            f << zip.read(entry)
          }
          @sheet_files[nr.to_i-1] = tmpdir+'/'+"roo_sheet#{nr}"
        end
        if entry.to_s.downcase =~ /comments([0-9]+).xml$/
          nr = $1
          open(tmpdir+'/'+"roo_comments#{nr}",'wb') {|f|
            f << zip.read(entry)
          }
          @comments_files[nr.to_i-1] = tmpdir+'/'+"roo_comments#{nr}"
        end
      }
    }
    # return
  end

  # extract files from the zip file
  def extract_content(tmpdir, zipfilename)
    Roo::ZipFile.open(@filename) do |zip|
      process_zipfile(tmpdir, zipfilename,zip)
    end
  end

  # extract single sheet and open it for parsing
  def extract_sheet_at_index(tmpdir, idx)
    Roo::ZipFile.open(@filename) do |zip|
      Roo::ZipFile.open(@filename) do |zf|
        zf.entries.each do |entry|
          next unless entry.to_s.downcase =~ /sheet([#{idx+1}]+).xml$/
          file_name = "roo_sheet#{idx}"
          open(tmpdir+'/'+file_name,'wb') {|f| f << zip.read(entry) }
          return File.join(tmpdir, file_name)
        end
      end
    end
  end

  # read the shared strings xml document
  def read_shared_strings(doc)
    doc.xpath("/xmlns:sst/xmlns:si").each do |si|
      shared_table_entry = ''
      si.children.each do |elem|
        if elem.name == 'r' and elem.children
          elem.children.each do |r_elem|
            if r_elem.name == 't'
              shared_table_entry << r_elem.content
            end
          end
        end
        if elem.name == 't'
          shared_table_entry = elem.content
        end
      end
      @shared_table << shared_table_entry
    end
  end

  # read the styles elements of an excelx document
  def read_styles(doc)
    @cellXfs = []

    @numFmts = Hash[doc.xpath("//xmlns:numFmt").map do |numFmt|
      [numFmt['numFmtId'], numFmt['formatCode']]
    end]
    fonts = doc.xpath("//xmlns:fonts/xmlns:font").map do |font_el|
      Font.new.tap do |font|
        font.bold = !font_el.xpath('./xmlns:b').empty?
        font.italic = !font_el.xpath('./xmlns:i').empty?
        font.underline = !font_el.xpath('./xmlns:u').empty?
      end
    end

    doc.xpath("//xmlns:cellXfs").each do |xfs|
      xfs.children.each do |xf|
        @cellXfs << xf['numFmtId']
        @style_definitions << fonts[xf['fontId'].to_i]
      end
    end
  end

  # convert internal excelx attribute to a format
  def attribute2format(s)
    id = @cellXfs[s.to_i]
    @numFmts[id] || Format::STANDARD_FORMATS[id.to_i]
  end

  def base_date
    @base_date ||= read_base_date
  end

  # Default to 1900 (minus one day due to excel quirk) but use 1904 if
  # it's set in the Workbook's workbookPr
  # http://msdn.microsoft.com/en-us/library/ff530155(v=office.12).aspx
  def read_base_date
    base_date = Date.new(1899,12,30)
    @workbook_doc.xpath("//xmlns:workbookPr").map do |workbookPr|
      if workbookPr["date1904"] && workbookPr["date1904"] =~ /true|1/i
        base_date = Date.new(1904,01,01)
      end
    end
    base_date
  end

end # class
