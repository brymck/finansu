require 'albacore'

task :default => :build

desc "Builds the application."
msbuild :build do |msb|
  msb.properties :configuration => :Release
  msb.targets :Clean, :Build
  msb.solution = "FinAnSu.sln"
end

# This requires that nunit-console.exe is in your %Path$
desc "NUnit Test Runner Example"
nunit :test do |nunit|
  nunit.command = "nunit-console.exe"
  nunit.assemblies "FinAnSu.Test/bin/Release/FinAnSu.Test.dll"
end
