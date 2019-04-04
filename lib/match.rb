class Match
  require 'roo'
  require 'neo4j/core/cypher_session/adaptors/http'
  def self.script()
  begin
    @prod_id_arr= 0, @id_arr=0, id_arr =[], @prod_id =0, @compare_name = 0, @is_present, val =0, compare_name = 0, str=""
    neo4j_adaptor = Neo4j::Core::CypherSession::Adaptors::HTTP.new('http://neo4j:graphdb@localhost:11008')
    @neo4j_session = Neo4j::ActiveBase.current_session = Neo4j::Core::CypherSession.new(neo4j_adaptor)
    file = Roo::Spreadsheet.open('/mnt/c/users/Manoj/Downloads/compare_tool_template.xlsx')
    file.sheet('Sheet1').each do |row|	
      if row[1] != 'Product ID'
      #prod_id_arr = row[1]				
	 if row[1] != nil
	   if(compare_name != 0)
	     mapping()
	   end
	   prod_id_arr 	  = row[1]
	   id_arr         = prod_id_arr.split(',')#split product ids
	   compare_name   = row[2]
	   @id_arr        = id_arr
	   @compare_name  = compare_name
	end#end if

	display_order        = row[3]
	option_heading       = row[4]
	option_sub_heading   = row[5]
	option_description   = row[6]
	display              = row[7]
	if prod_id_arr != nil
	   id_arr.each do |product_id|
	      prod_id = product_id
	      res1 = @neo4j_session.query("match (Product) where Product.prod_id = #{prod_id} return count(Product)")
	      @is_present = res1.rows[0].first
	      if @is_present == 0
	         str =  "CREATE (n:Product{prod_id:#{prod_id}, title:'#{compare_name}'})"
		 @neo4j_session.query(str)	  
	      end 
	   end
	end

	if display_order != nil
	   if @is_present == 0  #Check If product Already Present or not
	      if option_sub_heading != nil
		 str = "CREATE(n:ProductAttribute{name:'#{compare_name}',display_order:#{display_order}, option_heading:'#{option_heading}', option_sub_heading:'#{option_sub_heading}', option_description:'#{option_description}', display:'#{display}'})"
		 @neo4j_session.query(str)
	      else
		str = "CREATE(n:ProductAttribute{name:'#{compare_name}',display_order:#{display_order}, option_heading:'#{option_heading}', option_sub_heading:'#{option_sub_heading}', option_description:'#{option_description}', display:'#{display}'})"
		@neo4j_session.query(str)
	      end	      
	   end		
	end
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
        if @is_present == 0
	   str = "MATCH (a:Product),(b: ProductAttribute)WHERE a.prod_id = #{product_id} AND b.name = '#{@compare_name}' CREATE (a)-[r:HAS_PRODUCT_ATTRIBUTE]->(b) RETURN type(r)"
	   @neo4j_session.query(str)
	end
      end
   end#abc Method end
   Match.script()
 end#Class end
