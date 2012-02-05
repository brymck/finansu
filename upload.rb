#!/usr/bin/env ruby -Ku
PROJECT_NAME  = "FinAnSu"
REPO_NAME     = "brymck/finansu"

this_path     = File.expand_path(File.dirname(__FILE__))
upload_script = File.join(this_path, "Lib", "github_upload", "upload.rb")
zip_files     = Dir[File.join(this_path, PROJECT_NAME, "bin", "Release",
                              "#{PROJECT_NAME}-*.zip")]

zip_files.each do |zip_file|
  puts upload_script
  puts "Uploading #{zip_file}..."
  system "ruby \"%s\" \"%s\" %s" % [upload_script, zip_file, REPO_NAME]
end
