#! c:\perl\bin\perl.exe
#-------------------------------------------------------------------------
# Rip - RegRipper, CLI version
# Use this utility to run a plugins file or a single plugin against a Reg
# hive file.
# 
# Output goes to STDOUT
# Usage: see "_syntax()" function
#
# Change History
#   20090102 - updated code for relative path to plugins dir
#   20080419 - added '-g' switch (experimental)
#   20080412 - added '-c' switch
#
# copyright 2008 H. Carvey, keydet89@yahoo.com
#-------------------------------------------------------------------------
use strict;
use Parse::Win32Registry qw(:REG_);
use Getopt::Long;

# Included to permit compiling via Perl2Exe
#perl2exe_include "Parse/Win32Registry.pm";
#perl2exe_include "Parse/Win32Registry/Entry.pm";
#perl2exe_include "Parse/Win32Registry/Key.pm";
#perl2exe_include "Parse/Win32Registry/Value.pm";
#perl2exe_include "Parse/Win32Registry/File.pm";
#perl2exe_include "Parse/Win32Registry/Win95/File.pm";
#perl2exe_include "Parse/Win32Registry/Win95/Key.pm";
#perl2exe_include "Encode/Unicode.pm";

my %config;
Getopt::Long::Configure("prefix_pattern=(-|\/)");
GetOptions(\%config,qw(reg|r=s file|f=s csv|c guess|g plugin|p=s list|l help|?|h));

# Code updated 20090102
my @path;
my $str = $0;
($^O eq "MSWin32") ? (@path = split(/\\/,$0))
                   : (@path = split(/\//,$0));
$str =~ s/($path[scalar(@path) - 1])//;
my $plugindir = $str."plugins/";
#print "Plugins Dir = ".$plugindir."\n";
# End code update
my $VERSION = "20090102";

if ($config{help} || !%config) {
	_syntax();
	exit;
}

#-------------------------------------------------------------
# 
#-------------------------------------------------------------
if ($config{list}) {
	my @plugins;
	opendir(DIR,$plugindir) || die "Could not open $plugindir: $!\n";
	@plugins = readdir(DIR);
	closedir(DIR);

	my $count = 1; 
	print "Plugin,Version,Hive,Description\n" if ($config{csv});
	foreach my $p (@plugins) {
		next unless ($p =~ m/\.pl$/);
		my $pkg = (split(/\./,$p,2))[0];
		$p = $plugindir.$p;
		eval {
			require $p;
			my $hive    = $pkg->getHive();
			my $version = $pkg->getVersion();
			my $descr   = $pkg->getShortDescr();
			if ($config{csv}) {
				print $pkg.",".$version.",".$hive.",".$descr."\n";
			}
			else {
				print $count.". ".$pkg." v.".$version." [".$hive."]\n";
#				printf "%-20s %-10s %-10s\n",$pkg,$version,$hive;
				print  "   - ".$descr."\n\n";
				$count++;
			}
		};
		print "Error: $@\n" if ($@);
	}
	exit;
}

#-------------------------------------------------------------
# 
#-------------------------------------------------------------
if ($config{file}) {
# First, check that a hive file was identified, and that the path is
# correct
	my $hive = $config{reg};
	die "You must enter a hive file path/name.\n" if ($hive eq "");
	die $hive." not found.\n" unless (-e $hive);
	
	my %plugins = parsePluginsFile($config{file});
	if (%plugins) {
		logMsg("Parsed Plugins file.");
	}
	else {
		logMsg("Plugins file not parsed.");
		exit;
	}
	foreach my $i (sort {$a <=> $b} keys %plugins) {
		eval {
			require "plugins\\".$plugins{$i}."\.pl";
			$plugins{$i}->pluginmain($hive);
		};
		if ($@) {
			logMsg("Error in ".$plugins{$i}.": ".$@);
		}
		logMsg($plugins{$i}." complete.");
		rptMsg("-" x 40);
	}
}

#-------------------------------------------------------------
# 
#-------------------------------------------------------------
if ($config{reg} && $config{guess}) {
# Attempt to guess which kind of hive we have
	my $hive = $config{reg};
	die "You must enter a hive file path/name.\n" if ($hive eq "");
	die $hive." not found.\n" unless (-e $hive);
	
	my $reg;
	my $root_key;
	my %guess;
	eval {
		$reg = Parse::Win32Registry->new($hive);
	  $root_key = $reg->get_root_key;
	};
	::rptMsg($config{reg}>" may not be a valid hive.") if ($@);
	
# Check for SAM
	eval {
		$guess{sam} = 1 if (my $key = $root_key->get_subkey("SAM\\Domains\\Account\\Users"));
	};
# Check for Software	
	eval {
		$guess{software} = 1 if ($root_key->get_subkey("Microsoft\\Windows\\CurrentVersion") &&
				$root_key->get_subkey("Microsoft\\Windows NT\\CurrentVersion"));
	};

# Check for System	
	eval {
		$guess{system} = 1 if ($root_key->get_subkey("MountedDevices") &&
				$root_key->get_subkey("Select"));
	};
	
# Check for Security	
	eval {
		$guess{security} = 1 if ($root_key->get_subkey("Policy\\Accounts") &&
				$root_key->get_subkey("Policy\\PolAdtEv"));
	};
# Check for NTUSER.DAT	
	eval {
		$guess{ntuser} = 1 if ($root_key->get_subkey("Software\\Microsoft\\Windows\\CurrentVersion"));
		
	};	
	
	foreach my $g (keys %guess) {
		::rptMsg(sprintf "%-8s = %-2s",$g,$guess{$g});
	}
}

#-------------------------------------------------------------
# 
#-------------------------------------------------------------
if ($config{plugin}) {
# First, check that a hive file was identified, and that the path is
# correct
	my $hive = $config{reg};
	die "You must enter a hive file path/name.\n" if ($hive eq "");
	die $hive." not found.\n" unless (-e $hive);	
		
# check to see if the plugin exists
	my $plugin = $config{plugin};
	my $pluginfile = $plugindir.$config{plugin}."\.pl";
	die $pluginfile." not found.\n" unless (-e $pluginfile);
	
	eval {
		require $pluginfile;
		$plugin->pluginmain($hive);
	};
	if ($@) {
		logMsg("Error in ".$pluginfile.": ".$@);
	}	
}

sub _syntax {
	print<< "EOT";
Rip v.$VERSION - CLI RegRipper tool	
Rip [-r Reg hive file] [-f plugin file] [-p plugin module] [-l] [-h]
Parse Windows Registry files, using either a single module, or a plugins file.
All plugins must be located in the \"plugins\" directory; default plugins file 
used if no other filename given is \"plugins\\plugins\"\.

  -r Reg hive file...Registry hive file to parse
  -g ................Guess the hive file (experimental)
  -f [plugin file]...use the plugin file (default: plugins\\plugins)
  -p plugin module...use only this module
  -l ................list all plugins
  -c ................Output list in CSV format (use with -l)
  -h.................Help (print this information)
  
Ex: C:\\>rr -r c:\\case\\system -f system
    C:\\>rr -r c:\\case\\ntuser.dat -p userassist
    C:\\>rr -l -c

All output goes to STDOUT; use redirection (ie, > or >>) to output to a file\.
  
copyright 2008 H. Carvey
EOT
}

#-------------------------------------------------------------
# 
#-------------------------------------------------------------
sub logMsg {
	print STDERR $_[0]."\n";
}

#-------------------------------------------------------------
# 
#-------------------------------------------------------------
sub rptMsg {
	binmode STDOUT,":utf8";
	print $_[0]."\n";
}

#-------------------------------------------------------------
# parsePluginsFile()
# Parse the plugins file and get a list of plugins
#-------------------------------------------------------------
sub parsePluginsFile {
	my $file = $_[0];
	my %plugins;
# Parse a file containing a list of plugins
# Future versions of this tool may allow for the analyst to 
# choose different plugins files	
	my $pluginfile = $plugindir.$file;
	if (-e $pluginfile) {
		open(FH,"<",$pluginfile);
		my $count = 1;
		while(<FH>) {
			chomp;
			next if ($_ =~ m/^#/ || $_ =~ m/^\s+$/);
#			next unless ($_ =~ m/\.pl$/);
			next if ($_ eq "");
			$_ =~ s/^\s+//;
			$_ =~ s/\s+$//;
			$plugins{$count++} = $_; 
		}
		close(FH);
		return %plugins;
	}
	else {
		return undef;
	}
}

#-------------------------------------------------------------
# getTime()
# Translate FILETIME object (2 DWORDS) to Unix time, to be passed
# to gmtime() or localtime()
#-------------------------------------------------------------
sub getTime($$) {
	my $lo = shift;
	my $hi = shift;
	my $t;

	if ($lo == 0 && $hi == 0) {
		$t = 0;
	} else {
		$lo -= 0xd53e8000;
		$hi -= 0x019db1de;
		$t = int($hi*429.4967296 + $lo/1e7);
	};
	$t = 0 if ($t < 0);
	return $t;
}