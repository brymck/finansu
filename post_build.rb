#!/usr/bin/env ruby -Ku
require "fileutils"
require "zip/zip"

PROJECT_NAME = "FinAnSu"

# Writes a blank row then a header row
def write_header(text)
  puts
  puts "---- %s ----" % text
end

# Writes a status row, by necessity indenting at least once
def write_status(text, opts = {})
  opts = { :indent => 0 }.merge(opts)
  puts "  %s%s" % ["  " * opts[:indent], text]
end

# Write a series of statuses
def write_statuses(array_or_hash, opts = {})
  opts = { :indent => 0 }.merge(opts)

  if array_or_hash.is_a?(Hash)
    array_or_hash.each do |key, value|
      write_status "#{key}:", :indent => opts[:indent]
      write_status value, :indent => opts[:indent] + 1
    end
  elsif array_or_hash.is_a?(Array)
    array_or_hash.each do |value|
      write_status value, :indent => opts[:indent]
    end
  else
    write_status array_or_hash, :indent => opts[:indent]
  end
end

def last_directory(path)
  File.split(path).last
end

# Grabs the version number by parsing Properties\AssemblyInfo.cs
def version(path = ARGV[0])
  return @version if defined?(@version)

  File.open(File.join(path, PROJECT_NAME, "Properties", "AssemblyInfo.cs")) do |file|
    file.each_line do |line|
      if line =~ /AssemblyFileVersion\(\"([0-9.]+)\"\)/
        @version = $1.gsub(/(\.0)*$/, "")
        return @version
      end
    end
  end

  @version = nil
end

# Create an empty directory at the supplied +dir+ by making all folders along
# the path and then clearing its contents, if necessary
def create_empty_directory(path)
  FileUtils.mkdir_p path
  FileUtils.rm_rf Dir["#{path}/."], :secure => true
end

def copy_addin_files(to_path, xll_name)
  # Copy DLL
  FileUtils.cp File.join(@dirs[:release], "#{PROJECT_NAME}.dll"), to_path

  # Copy XLL add-in
  FileUtils.cp File.join(@dirs[:excel_dna], xll_name),
          File.join(to_path, "#{PROJECT_NAME}.xll")

  # Copy DNA files
  FileUtils.cp File.join(@dirs[:resources], "#{PROJECT_NAME}.dna"), to_path
end

# Build list of relevant directories
def build_directory_list(solution_dir)
  project_dir = File.join(solution_dir, PROJECT_NAME)

  @dirs = {
    :solution  => solution_dir,
    :project   => project_dir,
    :excel_dna => File.join(solution_dir, "Lib", "ExcelDna", "Distribution"),
    :release   => File.join(project_dir, "bin", "Release"),
    :x86       => File.join(project_dir, "bin", "Release", "x86"),
    :x64       => File.join(project_dir, "bin", "Release", "x64"),
    :resources => File.join(project_dir, "Resources"),
    :addins    => File.join(ENV["AppData"], "Microsoft", "AddIns")
  }
end

def pack_addin(path)
  Dir.chdir(path) do
    FileUtils.cp "#{PROJECT_NAME}.xll", "ExcelDna.xll"
    FileUtils.cp File.join(@dirs[:excel_dna], "ExcelDnaPack.exe"), "."
    FileUtils.cp File.join(@dirs[:excel_dna], "ExcelDna.Integration.dll"), "."
    puts %x[#{File.join(path, "ExcelDnaPack")} #{File.join(path, "FinAnSu.dna")} /y]
  end
end

def zip_addin(path)
  zip_name = "#{PROJECT_NAME}-%s_%s.zip" % [version, last_directory(path)]
  File.delete(zip_name) if File.exist?(zip_name)

  # Move latest FinAnSu to main release directory
  FileUtils.cp File.join(path, "FinAnSu-packed.xll"),
               File.join(@dirs[:release], "FinAnSu.xll")

  Dir.chdir(@dirs[:release]) do
    # Clear archive
    system "7za d -tzip #{zip_name}"

    # Add files to archive
    system "7za a -tzip #{zip_name} Examples.xls"
    system "7za a -tzip #{zip_name} FinAnSu.xll"
    system "7za a -tzip #{zip_name} install.bat"
    system "7za a -tzip #{zip_name} Readme.txt"
  end
end

def copy_to_addin_directory(path)
  FileUtils.cp File.join(path, "FinAnSu-packed.xll"),
               File.join(@dirs[:addins], "FinAnSu_#{last_directory(path)}.xll")
end

write_status "Version #{version}"

# Build directory list
write_header "Building directory list"
@dirs = {}
build_directory_list ARGV[0]
write_statuses @dirs

write_header "Creating empty directories for add-in versions"
create_empty_directory @dirs[:x86]
create_empty_directory @dirs[:x64]
write_statuses Dir["*/"].map(&:chop)

write_header "Copying add-in files"
copy_addin_files @dirs[:x86], "ExcelDna.xll"
copy_addin_files @dirs[:x64], "ExcelDna64.xll"
write_statuses Dir["*/*"]

write_header "Packing add-ins"
pack_addin @dirs[:x86]
pack_addin @dirs[:x64]
write_statuses Dir["*/*-packed.xll"]

write_header "Zipping add-ins"
zip_addin @dirs[:x86]
zip_addin @dirs[:x64]
Dir["#{PROJECT_NAME}-*.zip"].each do |zip_file|
  write_status zip_file
  write_statuses Zip::ZipFile.open(zip_file).entries, :indent => 1
end

write_header "Copying add-ins to add-in directory"
copy_to_addin_directory @dirs[:x86]
copy_to_addin_directory @dirs[:x64]

# Write new version number
write_header "Writing version number to text file"
puts version_path = File.join(@dirs[:solution], "VERSION")
File.open(version_path, "w") { |f| f.puts version }
write_status File.read(version_path)
