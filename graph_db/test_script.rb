#require_relative 'bonzo_db_util_1.rb'
require 'roo'

begin
	prod_id_arr= [],b=[], id_arr =[], prod_id =0, compare_name = 0, val =0 
#	mw_connection = getMWConnection()


	file = Roo::Spreadsheet.open('/mnt/c/users/Manoj/Downloads/compare_tool_template.xlsx')
	
	file.sheet('Sheet1').each do |row|	
	#	p row[0..7]
		
#		if row[1] != nil
			if row[1] != 'Product ID'
			#	p row[0..7]
				prod_id_arr = row[1]
				
				if row[1] != nil
					id_arr = prod_id_arr.split(',')#split product ids
				end

				id_arr.each do |product_id|
					prod_id              = product_id
					compare_name         = row[2]
					display_order        = row[3]
					option_heading       = row[4]
					option_sub_heading   = row[5]
					option_description   = row[6]
					display              = row[7]
					
					if prod_id_arr != nil
					#	result  = mw_connection.execute("Insert into compare_products (product_id,compare_data_name) values('#{prod_id}',
					#			       '#{compare_name}')")
					#	puts prod_id_arr
						puts "CREATE (n:Product{prod_id:#{prod_id}, title:'Product-#{prod_id}'})"
					end
					
					if display_order != nil
					#	result1 = mw_connection.execute("INSERT INTO mw_integration.compare_product_details
					#				(product_id, display_order, option_heading, option_sub_heading,
					#				 option_discription, display)
					#				 VALUES('#{prod_id}', '#{display_order}','#{option_heading}', '#{option_sub_heading}', 
					#				'#{option_description}', '#{display}')")
					end
				end
			end
	#	end





	end
        rescue => exception
	puts exception
	ensure
	puts "Executing ensure"
	closeMWConnection(mw_connection)
end
