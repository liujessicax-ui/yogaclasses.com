require 'webrick'
dir = File.dirname(File.expand_path(__FILE__))
server = WEBrick::HTTPServer.new(Port: 8080, DocumentRoot: dir)
trap('INT') { server.shutdown }
server.start
