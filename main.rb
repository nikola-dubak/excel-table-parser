require_relative './Parser'

tables = Parser.read_excel_tables('./test.xlsx')

# Print table matrices
tables.each do |table|
    puts table.sheet_name + ':'
    puts table.get_matrix().inspect
end
puts

# Test enumerable
print 'Cells: '
tables.first.each do |cell|
    print cell.to_s + ', '
end
puts

# Test row accessor
print 'Second row: '
puts tables.first.row(1).inspect
puts

# Test column
print 'Indeks column: '
puts tables.first['Indeks'].inspect
puts

# Test column methods
print 'Indeks column method: '
puts tables.first.Indeks.inspect
puts

# Test column row method
print 'Indeks column row method: '
puts tables.first.Indeks.rn1.inspect
puts

# Test table addition
print 'Added tables: '
tables[0] = tables[0] + tables[1]
puts tables[0].get_matrix.inspect
puts

# Test table subtraction
print 'Subtracted tables: '
tables[0] = tables[0] - tables[1]
puts tables[0].get_matrix.inspect
