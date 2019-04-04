class Graphdb
  require 'roo'
  require 'fileutils'
  require 'neo4j/core/cypher_session/adaptors/http'
  def self.script()
  begin
    p "Start"
    @prod_id_arr= 0, @id_arr=0, id_arr =[], @prod_id =0, @compare_name = 0, @is_present, val =0, compare_name = 0, str=""
    @reg = /^[0-9]+$/
    
    @neo4j_session = $start_neo4j_session
    
    file = Roo::Spreadsheet.open('/mnt/c/users/Manoj/Downloads/compare_tool_template_1.xlsx')
    file.sheet('Sheet1').each do |row|	
      if row[1] != 'Product ID'
      #prod_id_arr = row[1]				
	 if row[1] != nil
	   if(compare_name != 0)
	     mapping()
	   end
	   prod_id_arr 	  = row[1].to_s.gsub(/ /, "").gsub(/\n/,"")
	   id_arr         = prod_id_arr.split(',')#split product ids
	   compare_name   = row[2].to_s.gsub(/'/, "")
	   @id_arr        = id_arr
	   @compare_name  = compare_name
	 end#end if

	 display_order        = row[3]
	 option_heading       = row[4].to_s.gsub(/'/, "")
	 option_sub_heading   = row[5].to_s.gsub(/'/, "")
	 option_description   = row[6].to_s.gsub(/'/, "")
	 display              = row[7].to_s.gsub(/'/, "")
	
	if prod_id_arr != nil
	      compare_prod_is_present = @neo4j_session.query("match(n:ProductAttribute{name:'#{compare_name}'})--(Product) return distinct(Product.prod_id)")
	      product_id_arr = compare_prod_is_present.rows.flatten
	      array_length = product_id_arr.length
              p prod_id_arr
	      if array_length > 0
		  @neo4j_session.query("match(n:Product) where n.prod_id in #{product_id_arr} detach delete n")
	      end	      
	      
	      prod_attribute_count = @neo4j_session.query("MATCH(n:ProductAttribute) WHERE n.name = '#{compare_name}' return count(n)")
	      prod_attribute_count = prod_attribute_count.rows[0].first
	      if prod_attribute_count != 0
		  @neo4j_session.query("match(n:ProductAttribute) where n.name='#{compare_name}' delete n")
	      end 

	      id_arr.each do |product_id|
	       prod_id = product_id
	       if @reg.match(prod_id)  #check regular expression
	         result = @neo4j_session.query("match (Product) where Product.prod_id = #{prod_id} return count(Product)")
	         @is_present = result.rows[0].first
	         #if @is_present == 0
	            create_product_node =  "CREATE (n:Product{prod_id:#{prod_id}, title:'#{compare_name}'})"
		    str_resp = @neo4j_session.query(create_product_node)
	         #end
	       else
		       p "product_id is #{prod_id}"
	       end
	   end
	end
        
	if display_order != nil
	   prod_attribute_count =  @neo4j_session.query("match (n:ProductAttribute) where n.name='#{compare_name}' and n.display_order=#{display_order} and n.option_description='#{option_description}' return count(n)") 
              prod_attribute_count = prod_attribute_count.rows[0].first
	      #p prod_attribute_count

	      if (prod_attribute_count == 0)
		 create_prod_attribute = "CREATE(n:ProductAttribute{name:'#{compare_name}',display_order:#{display_order}, option_heading:'#{option_heading}', option_sub_heading:'#{option_sub_heading}', option_description:'#{option_description}', display:'#{display}'})"  
		@neo4j_session.query(create_prod_attribute)
	      else
		#p "Update attribute" 
		update_prod_attribute = "match (n:ProductAttribute) where n.name='#{compare_name}' and n.display_order=#{display_order} and n.option_description='#{option_description}' set n.display='#{display}'"
		str_resp = neo4j_session.query(update_prod_attribute)
	      end

	end#Check display order nil end
       
       end
      end#Do end
      if(@id_arr != 0)
	mapping()
      end
      rescue => exception
	puts exception
      ensure
	puts "Executing ensure"
      end#Begin end
    end#Method end
    def self.mapping()
	@id_arr.each do |product_id|
	  if @reg.match(product_id)
	     str = "MATCH (a:Product),(b: ProductAttribute) WHERE a.prod_id = #{product_id} AND b.name = '#{@compare_name}' merge(a)-[r:HAS_COMPARE_ATTRIBUTE]->(b) RETURN type(r)"
	     @neo4j_session.query(str)
	  end
      end
   end#abc Method end
   Graphdb.script()
end#Class end
