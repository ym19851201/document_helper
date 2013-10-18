require "bundler"
Bundler.require

class BookHelper < Spreadsheet::Workbook
#   def create_worksheet opts = {}
#     opts[:name] ||= client("Worksheet#{@worksheets.size.next}", 'UTF-8')
#     add_worksheet SheetHelper.new(opts)
#   end

#   def worksheet idx
#     case idx
#     when Integer
#       @worksheets[idx]
#     when String
#       @worksheets.find do |sheet| sheet.name == idx end
#     end
#   end
end

class Spreadsheet::Worksheet
  def [](*args)
    if args.length == 2 && args[0].class == Fixnum && args[1].class == Fixnum
#       super
      row(args[0])[args[1]]
    elsif args.length == 1 && args[0].class == String
      range = args[0].split(':')
      y1, x1 = convert_to_xy(range[0].upcase)

      if range[1]
        y2, x2 = convert_to_xy(range[1].upcase)
        values = []
        for y in y1..y2
          x_values = []
          for x in x1..x2
#             x_values << super(y, x)
            x_values << row(y)[x]
          end
          values << x_values
        end
        return values
      end
      row(y1)[x1]
#       super(y, x)
    end
  end

  def []=(*args)
    if args.length == 3 && args[0].class == Fixnum && args[1].class == Fixnum
#       super
      row(args[0])[args[1]] = args[2]
    elsif args.length == 2 && args[0].class == String
      range = args[0].split(':')
      y1, x1 = convert_to_xy(range[0].upcase)
      if range[1]
        y2, x2 = convert_to_xy(range[1].upcase)
        values = args[1]
        values.each_with_index do |value, y|
          value.each_with_index do |v, x|
#             super(y1 + y, x1 + x, v) if y1 + y <= y2 && x1 + x <= x2
            row(y1 + y)[x1 + x]=v if y1 + y <= y2 && x1 + x <= x2
          end
        end
      else
#         super(y1, x1, args[1])
        row(y1)[x1] = args[1]
      end
    end
  end

  private
  def convert_to_xy(str)
    if str =~ /(\w+)(\d+)/
      x = $1
      y = $2.to_i
      x = al_to_num(x, x.length - 1)
      return y - 1, x - 1
    end
  end

  def al_to_num(str, i, num_val=0)
    index = str.length - 1 - i
    num_val += (((str[index].ord - 65)) + 1) * (26 ** i)
    if i == 0
      return num_val 
    end
    al_to_num(str, i - 1, num_val)
  end
end

