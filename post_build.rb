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
    :project   => project_dir,
    :excel_dna => File.join(solution_dir, "Lib", "ExcelDna", "Distribution"),
    :release   => File.join(project_dir, "bin", "Release"),
    :x86       => File.join(project_dir, "bin", "Release", "x86"),
    :x64       => File.join(project_dir, "bin", "Release", "x64"),
    :resources => File.join(project_dir, "Resources")
  }
end

def pack_addin(path)
  Dir.chdir(path) do
    FileUtils.cp "#{PROJECT_NAME}.xll", "ExcelDna.xll"
    FileUtils.cp File.join(@dirs[:excel_dna], "ExcelDnaPack.exe"), "."
    FileUtils.cp File.join(@dirs[:excel_dna], "ExcelDna.Integration.dll"), "."
    system "ExcelDnaPack FinAnSu.dna /y"
  end
end

def zip_addin(path)
  suffix = File.split(path).last
  zip_name = "#{PROJECT_NAME}-%s_%s.zip" % [version, suffix]
  File.delete(zip_name) if File.exist?(zip_name)

  Dir.chdir(@dirs[:release]) do
    Zip::ZipFile.open(zip_name, Zip::ZipFile::CREATE) do |zip_file|
      zip_file.add "Examples.xls", "Examples.xls"
      zip_file.add "install.bat", "install.bat"
      zip_file.add "README.txt", "README.txt"
      zip_file.add "#{PROJECT_NAME}.xll", File.join(path, "#{PROJECT_NAME}-packed.xll")
    end
  end
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
