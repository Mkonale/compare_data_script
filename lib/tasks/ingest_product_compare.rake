namespace :ingest_product_compare do
  desc "Process products from FTP to Neo4J"
  #ingest_product_compare:process_products_to_neo4j
  task process_products_to_neo4j: :environment do
    require 'neo4j/core/cypher_session/adaptors/http'
    
   neo4j_adaptor = Neo4j::Core::CypherSession::Adaptors::HTTP.new('http://34.229.231.48:7474')
   #neo4j_adaptor = Neo4j::Core::CypherSession::Adaptors::HTTP.new('http://neo4j:graphdb@localhost:11008')
    
    $start_neo4j_session = Neo4j::ActiveBase.current_session = Neo4j::Core::CypherSession.new(neo4j_adaptor)
    
    #require './lib/match'
    require './lib/graphdb'

    puts "############################################"
  end
end
