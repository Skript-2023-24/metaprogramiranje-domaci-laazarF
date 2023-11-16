require "google_drive"
require 'google/apis/sheets_v4'

session = GoogleDrive::Session.from_config("config.json")
spreadsheet_id = "1T7NEci8BC1E3tg6QOXmf2V4000FDGXvM8Ybk6uFHlmc"

ws = session.spreadsheet_by_key(spreadsheet_id).worksheets[0]

class SheetTable
  include Enumerable

  def initialize(sheet_data, spreadsheet_id, session)
    @data = sheet_data
    @spreadsheet_id = spreadsheet_id
    @session = session
  end

  def [](column_name)
    header = @data.rows.first
    col_index = header.index(column_name)
    @data.rows.map { |row| row[col_index] }
  end

  def []=(column_name, row_index, value)
    header = @data.rows.first
    col_index = header.index(column_name)

    @data[row_index + 1, col_index + 1] = value
    @data.save
  end

  def row(index)
    @data.rows[index - 1]
  end

  def to_array
    @data.rows
  end

  def contains_total_or_subtotal?(row) #7.
    row.any? { |cell| cell.to_s.downcase.include?("total") || cell.to_s.downcase.include?("subtotal") }
  end

  def each
    @data.rows.each do |row|
      next if contains_total_or_subtotal?(row)

      row.each do |cell|
        yield cell
      end
    end
  end

  def headers
    @data.rows.first
  end

  def rows
    @data.rows.drop(1)
  end

  def +(other)
    raise 'Headers do not match' unless headers == other.headers

    combined_rows = (rows + other.rows).uniq
    SheetTable.new_sheet_with_combined_rows(@session, @spreadsheet_id, "Kombinovana+Tabela", headers, combined_rows)
  end

  def -(other)
    raise 'Headers do not match' unless headers == other.headers

    unique_rows = rows.reject do |row_t2|
      other.rows.any? { |row_t1| row_t2 == row_t1 }
    end
    SheetTable.new_sheet_with_combined_rows(@session, @spreadsheet_id, "Kombinovana-Tabela", headers, unique_rows)
  end

  def self.new_sheet_with_combined_rows(session, spreadsheet_id, new_sheet_title, headers, combined_rows)
    spreadsheet = session.spreadsheet_by_key(spreadsheet_id)

    new_ws = spreadsheet.add_worksheet(new_sheet_title)
    new_ws.update_cells(1, 1, [headers])
    new_ws.update_cells(2, 1, combined_rows)

    new_ws.save
  end

end

sheet_table = SheetTable.new(ws, spreadsheet_id, session)

array_data = sheet_table.to_array #1.

row_data = sheet_table.row(1) #2.

sheet_table.each do |cell| #3.
  #puts cell
end

column_data = sheet_table["Druga Kolona"] #5.
column_data.each do |cell|
  #puts cell
end

cell_data = sheet_table["Treca kolona"][2] #5.
#puts cell_data

sheet_table["Prva Kolona", 3] = 2557 #5.

ws1 = session.spreadsheet_by_key(spreadsheet_id).worksheets[1]
sheet_table1 = SheetTable.new(ws1, spreadsheet_id, session)

combined_table = sheet_table + sheet_table1 #8.
combined_table1 = sheet_table - sheet_table1 #9.


