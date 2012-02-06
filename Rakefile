require 'albacore'

PROJECT_NAME  = "FinAnSu"
REPO_NAME     = "brymck/finansu"

task :default => :build
task :release => [:build, :upload]

desc "Builds the application."
msbuild :build do |msb|
  msb.properties :configuration => :Release
  msb.targets :Clean, :Build
  msb.solution = "FinAnSu.sln"
end

# This requires that nunit-console.exe is in your %Path$
desc "Run NUnit tests"
nunit :test do |nunit|
  nunit.command = "nunit-console.exe"
  nunit.assemblies "FinAnSu.Test/bin/Release/FinAnSu.Test.dll"
end

desc "The current version"
task :version do
  puts File.read("VERSION")
end

desc "Uploads the latest zip files to GitHub"
task :upload do
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
end
