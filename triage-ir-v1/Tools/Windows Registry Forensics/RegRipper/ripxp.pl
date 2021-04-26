#! c:\perl\bin\perl.exe
#-------------------------------------------------------------------------
# RipXP - RegRipper, CLI version
# Use this utility to run plugin against a Reg hive file, as well as the
# corresponding hive files in the System Restore Points
# 
# Change history:
#   20090818 - Added check for hive type to allow plugins for "All"
#              hives (e.g., findexes.pl) to be run
#
# Output goes to STDOUT
# Usage: see "_syntax()" function
#
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
GetOptions(\%config,qw(reg|r=s dir|d=s guess|g plugin|p=s list|l help|?|h));

my $plugindir = "plugins\\";
my $VERSION = "20090818";
my @sids;

my %rp_files = ("System"      => "_REGISTRY_MACHINE_SYSTEM",
                "Software"    => "_REGISTRY_MACHINE_SOFTWARE",
                "SAM"         => "_REGISTRY_MACHINE_SAM",
                "Security"    => "_REGISTRY_MACHINE_SECURITY",
                "NTUSER\.DAT" => "_REGISTRY_USER_NTUSER_");

if ($config{help} || !%config) {
	_syntax();
	exit;
}

die "You must enter a hive file.\n" if (!$config{reg} && !$config{list});

# populate rest of global variables
\guessHive() if ($config{reg});

#-------------------------------------------------------------
# 
#-------------------------------------------------------------
if ($config{list}) {
	my @plugins;
	opendir(DIR,$plugindir) || die "Could not open $plugindir: $!\n";
	@plugins = grep {/\.pl$/}readdir(DIR);
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
# This is where the work is done
#-------------------------------------------------------------
if ($config{reg} && $config{dir} && $config{plugin}) {
# At this point, we should already have determined the type of the hive file, and
# obtained the SID for an NTUSER.DAT file via guessHive()
	
# Print some basic info first
	::rptMsg("RipXP v.".$VERSION);
	my $t = time;
	::rptMsg("Launched ".gmtime($t)." Z");
	::rptMsg("");
# Check the plugin to ensure that it is meant for the hive type
	my $plugin = $plugindir.$config{plugin}."\.pl";
	my $pkg = $config{plugin};
	my $hive;
	eval {
		require $plugin;
		$hive    = $pkg->getHive(); 
	};
	
	die $config{plugin}." plugin is not meant for ".$config{hive}." hives.\n" 
	  unless ($hive eq $config{hive} || $hive eq "All");
	::rptMsg($config{reg});
# Now, run the plugin against the hive file
	eval {
			$pkg->pluginmain($config{reg});
	};
	if ($@) {
		logMsg("Error in ".$plugin.": ".$@);
	}
		
# Get info on RP dirs
	my %rp_dirs;
	$config{dir} = $config{dir}."\\" unless ($config{dir} =~ m/\\$/);
	die $config{dir}." not found.\n" unless (-d $config{dir} && -e $config{dir});
	my @dirs;
	opendir(DIR,$config{dir}) || die "Could not access ".$config{dir}.": $!\n";
	@dirs = grep {/^RP/ && -d $config{dir}.$_."\\snapshot"} readdir(DIR);
	close(DIR);
	
	map{$rp_dirs{$_} = 1}@dirs;
	
# Cycle through each of the RP dirs
	foreach my $d (sort keys %rp_dirs) {
		my $snapshot_dir = $config{dir}.$d."\\snapshot\\";
		my $rp_log_file  = $config{dir}.$d."\\rp\.log";
		
		my $rp_file;
		if ($config{hive} eq "NTUSER\.DAT") {
			$rp_file = $rp_files{$config{hive}}.$config{sid};
		}
		else {
			$rp_file = $rp_files{$config{hive}};
		}
		::rptMsg("-" x 40);
# attempt to get some Restore Point information
		if (-e $rp_log_file) {
			my ($rp_time,$rp_type,$rp_str) = ::getRpData($rp_log_file);
			::rptMsg("Restore Point Info");
			::rptMsg("Description   : ".$rp_str);
			::rptMsg("Type          : ".$rp_type);
			::rptMsg("Creation Time : ".$rp_time);
		}
		::rptMsg("");
		::rptMsg($snapshot_dir.$rp_file);
		::rptMsg("");
		eval {
			$pkg->pluginmain($snapshot_dir.$rp_file);
		};
		if ($@) {
			logMsg("Error in ".$plugin.": ".$@);
		}	
	}
}

#-------------------------------------------------------------
# 
#-------------------------------------------------------------
sub _syntax {
	print<< "EOT";
RipXP v.$VERSION - CLI RegRipper tool	
RipXP [-r Reg hive file] [-p plugin module][-d RP dir][-lgh]
Parse Windows Registry files, using either a single module from the plugins folder.
Then parse all corresponding hive files from the XP Restore Points (extracted from
image) using the same plugin.

  -r Reg hive file...Registry hive file to parse
  -g ................Guess the hive file (experimental)
  -d RP directory....Path to the Restore Point directory
  -p plugin module...use only this module
  -l ................list all plugins
  -h.................Help (print this information)
  
Ex: C:\\>rip -g
    C:\\>rip -r d:\\cases\\ntuser.dat -d d:\\cases\\svi -p userassist

All output goes to STDOUT; use redirection (ie, > or >>) to output to a file\.
  
copyright 2008 H. Carvey
EOT
}

#-------------------------------------------------------------
# 
#-------------------------------------------------------------
sub logMsg {
#	print STDERR $_[0]."\n";
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
	my $file = shift || "plugins";
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

#-------------------------------------------------------------
# guessHive() - attempts to determine the hive type; if NTUSER.DAT,
#   attempt to retrieve the SID for the user; this function populates
#   global variables (%config, @sids)
#-------------------------------------------------------------
sub guessHive {
	my $reg;
	my $root_key;
	my $guess = 0;
	eval {
		$reg = Parse::Win32Registry->new($config{reg});
	  $root_key = $reg->get_root_key;
	};
	::rptMsg($config{reg}." may not be a valid hive.") if ($@);
	
# Check for SAM
	eval {
		if (my $key = $root_key->get_subkey("SAM\\Domains\\Account\\Users")) {
			$config{hive} = "SAM";
			$guess++;
		}
	};
# Check for Software	
	eval {
		if ($root_key->get_subkey("Microsoft\\Windows\\CurrentVersion") &&
				$root_key->get_subkey("Microsoft\\Windows NT\\CurrentVersion")) {
			$config{hive} = "Software";
			$guess++;
		}
	};

# Check for System	
	eval {
		if ($root_key->get_subkey("MountedDevices") && $root_key->get_subkey("Select")) {
			$config{hive} = "System";
			$guess++;
		}
	};
	
# Check for Security	
	eval {
		if ($root_key->get_subkey("Policy\\Accounts") &&	$root_key->get_subkey("Policy\\PolAdtEv")) {
			$config{hive} = "Security";
			$guess++;
		}
	};
# Check for NTUSER.DAT	
	eval {
	 	if ($root_key->get_subkey("Software\\Microsoft\\Windows\\CurrentVersion")) { 
	 		$config{hive} = "NTUSER\.DAT";
	 		$guess++;
	 		
	 		eval {
	 			my @subkeys = $root_key->get_subkey("Software\\Microsoft\\Protected Storage System Provider")->get_list_of_subkeys();
	 			map{push(@sids,$_->get_name())}@subkeys;
	 		};
	 		die "Error attempting to locate SID: $!\n" if ($@);
	 		$config{sid} = $sids[0] if (scalar(@sids) == 1);
	 	}
		
	};	
	
	die "Unable to determine hive file type for ".$config{reg}."\n" if ($guess < 1);
	
	if ($config{guess}) {
		::rptMsg("Guess    = ".$config{hive}) if ($guess == 1);
		if ($config{hive} eq "NTUSER\.DAT" && scalar(@sids) == 1) {
			::rptMsg("User SID = ".$config{sid});
		}
	}
}

#-------------------------------------------------------------
# getRpData()
#
#-------------------------------------------------------------
sub getRpData {
	my $file = $_[0];
	my $data;
	
	my %types = (0 => "Application Install",
		        1 => "Application Uninstall",
		        7 => "System CheckPoint",
		        10 => "Device Driver Install",
		        11 => "System Checkpoint",
		        12 => "Modify Settings",
		        13 => "Cancelled Operation");
		        
	my ($rp_time,$rp_type,$rp_str);
	
	eval {
		open(FH,"<",$file);
		binmode(FH);
		seek(FH,0x210,0);
		read(FH,$data,8);
		my @t = unpack("VV",$data);
		$rp_time = gmtime(::getTime($t[0],$t[1]));
				
		seek(FH,0x04,0);
		read(FH,$data,4);
		$rp_type = unpack("V",$data);
	
		my $offset = 0x10;
		my $tag = 1;
		my @strs;
		while($tag) {
			seek(FH,$offset,0);
			read(FH,$data,2);
			if (unpack("v",$data) == 0) {
				$tag = 0;
			}
			else {
				push(@strs,$data);
			}
			$offset += 2;
		}
		$rp_str = join('',@strs);
		$rp_str =~ s/\00//g;

	};
	return ($rp_time,$types{$rp_type},$rp_str);
}