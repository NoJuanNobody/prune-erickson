Gem::Specification.new do |s|
  s.name        = 'prune-erickson'
  s.version     = '0.0.1'
  s.date        = '2015-09-28'
  s.summary     = "prune-erickson"
  s.description = "An excel processing script to remove erroneous data"
  s.authors     = ["Alejandro Londono"]
  s.email       = 'alejandro.londono@erickson.com'
  s.files       = ["lib/prune-erickson.rb"]
  s.license       = 'apache'
  s.executables << 'prune-erickson'
  s.add_development_dependency "spreadsheet",  '~> 1.0', '>= 1.0.7'
end