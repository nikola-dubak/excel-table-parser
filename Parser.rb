require 'roo'

class Parser
    def self.read_excel_tables(file_path)
        excel_file = Roo::Excelx.new(file_path, {:expand_merged_ranges => true})
        tables = Array.new
        excel_file.each_with_pagename do |name, sheet|
            table = read_table_from_sheet(name, sheet)
            if (table != nil)
                tables.append(table)
            end
        end
        return tables
    end

    def self.read_table_from_sheet(name, sheet)
        columns = Array.new
        sheet.each_row_streaming(max_rows: 0) do |header|
            header.each do |header_column|
                columns.append(Column.new(header_column.value))
            end
        end
        if (columns.length == 0)
            return nil
        end

        sheet.each_row_streaming().drop(1).each do |row|
            column = 0
            if (row.any? { |cell| cell.value.to_s.casecmp('total').zero? or cell.value.to_s.casecmp('subtotal').zero? })
                next
            end
            row.each do |cell|
                if cell.to_s == nil
                    columns[column].add_value(sheet.cell(*cell.coordinate))
                else 
                    columns[column].add_value(cell.value)
                end
                column = column + 1
            end
        end
        return Table.new(name, columns)
    end
end

class Table include Enumerable
    attr_accessor :sheet_name, :columns
    def initialize(sheet_name, columns)
        @sheet_name = sheet_name
        @columns = columns
        @columns.each do |column|
            column.table = self
        end
        define_column_methods()
    end

    def get_matrix
        matrix = Array.new
        columns.each do |column|
            matrix.append(column.values)
        end
        return matrix.transpose()
    end

    def row(index)
        return get_matrix()[index];
    end

    def each(&block)
        matrix = get_matrix()
        matrix.each do |row|
            row.each do |cell|
                block.call(cell)
            end
        end
    end

    def [](column_name)
        column_values = nil
        @columns.each do |column|
            if (column.name == column_name)
                column_values = column.values
                break
            end
        end
        return column_values
    end

    def define_column_methods
        @columns.each do |column|
            define_singleton_method(column.name.gsub(/[^a-zA-z0-9]/, '_')) do
                return column
            end
        end
    end

    def +(table)
        if (@columns.count != table.columns.count)
            return nil
        end
        for i in 0...@columns.count
            if (@columns[i].name != table.columns[i].name)
                return nil
            end
        end
        for i in 0...@columns.count
            @columns[i].values.concat(table.columns[i].values)
        end
        return self
    end

    def -(table)
        if (@columns.count != table.columns.count)
            return nil
        end
        for i in 0...@columns.count
            if (@columns[i].name != table.columns[i].name)
                return nil
            end
        end

        self_matrix = get_matrix()
        table_matrix = table.get_matrix()
        rows_to_remove = Array.new
        for i in 0...@columns[0].values.count
            for j in 0...table.columns[0].values.count
                rows_to_remove.append(i) unless self_matrix[i] != table_matrix[j]
            end
        end

        rows_to_remove.uniq!
        rows_to_remove.sort!.reverse!
        rows_to_remove.each do |row_index|
            @columns.each do |column|
                column.values.delete_at(row_index)
            end
        end
        return self
    end
end

class Column include Enumerable
    attr_accessor :name, :values, :table
    def initialize(name)
        @name = name
        @values = Array.new
    end

    def add_value(value)
        @values.append(value)
        define_row_method(value)
    end

    def [](index)
        return values[index]
    end

    def sum
        total = 0
        values.each do |value|
            total += value
        end
        return total
    end

    def define_row_method(value)
        define_singleton_method(value.to_s) do
            return table.row(values.index(value))
        end
    end
end
