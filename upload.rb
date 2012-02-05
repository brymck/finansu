#!/usr/bin/env ruby -Ku
puts "This upload script doesn't work at the moment."
puts "Upload everything via https://github.com/brymck/finansu/downloads"
exit 0

PROJECT_NAME  = "FinAnSu"
REPO_NAME     = "brymck/finansu"

this_path     = File.expand_path(File.dirname(__FILE__))
upload_script = File.join(this_path, "Lib", "github_upload", "upload.rb")
zip_files     = Dir[File.join(this_path, PROJECT_NAME, "bin", "Release",
                              "#{PROJECT_NAME}-*.zip")]

zip_files.each do |zip_file|
  # Get the basename and build a description
  basename = File.basename(zip_file)
  if basename =~ /x86/
    description = "The latest 32-bit version"
  elsif basename =~ /x64/
    description = "The latest 64-bit version"
  end
  puts "Uploading #{basename}..."

  # Upload the files
  Dir.chdir(File.dirname(zip_file)) do
    system "ruby \"%s\" \"%s\" %s \"%s\"" % [upload_script, basename, REPO_NAME, description]
  end
end
