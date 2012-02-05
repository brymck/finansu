#!/usr/bin/env ruby -Ku
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

# {"expirationdate":"2112-02-05T09:32:38.000Z","signature":"HfasCLOJ1CWIRvbeXaD3wLseozk=","prefix":"downloads/brymck/finansu","redirect":false,"acl":"public-read","mime_type":"application/zip","accesskeyid":"1DWESVTPGHQVTX38V182","path":"downloads/brymck/finansu/FinAnSu-0.9.4_x86.zip","bucket":"github","policy":"ewogICAgJ2V4cGlyYXRpb24nOiAnMjExMi0wMi0wNVQwOTozMjozOC4wMDBaJywKICAgICdjb25kaXRpb25zJzogWwogICAgICAgIHsnYnVja2V0JzogJ2dpdGh1Yid9LAogICAgICAgIHsna2V5JzogJ2Rvd25sb2Fkcy9icnltY2svZmluYW5zdS9GaW5BblN1LTAuOS40X3g4Ni56aXAnfSwKICAgICAgICB7J2FjbCc6ICdwdWJsaWMtcmVhZCd9LAogICAgICAgIHsnc3VjY2Vzc19hY3Rpb25fc3RhdHVzJzogJzIwMSd9LAogICAgICAgIFsnc3RhcnRzLXdpdGgnLCAnJEZpbGVuYW1lJywgJyddLAogICAgICAgIFsnc3RhcnRzLXdpdGgnLCAnJENvbnRlbnQtVHlwZScsICcnXQogICAgXQp9"}
