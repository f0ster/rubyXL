module RubyXL
  class Parser

    # Parse <tt>.xslx</tt> file by reading it from local disk.
    def self.parse(src_file_path)
      begin
        ::RubyZip::File.open(src_file_path) { |zip_file|
          root = RubyXL::WorkbookRoot.parse_zip_file(zip_file)
          root.source_file_path = src_file_path
          root.workbook
        }
      rescue ::RubyZip::Error => e
        raise e, "XLSX file format error: #{e}", e.backtrace
      end
    end

    # Parse <tt>.xslx</tt> file contained in a stream (useful for receiving over HTTP).
    def self.parse_buffer(buffer)
      root = nil # RubyZip::File.open_buffer somehow fails to return the value from the block :(
      begin
        ::RubyZip::File.open_buffer(buffer) { |zip_file| root = RubyXL::WorkbookRoot.parse_zip_file(zip_file)  }
        root.workbook
      rescue ::RubyZip::Error => e
        raise e, "XLSX file format error: #{e}", e.backtrace
      end
    end

  end
end
