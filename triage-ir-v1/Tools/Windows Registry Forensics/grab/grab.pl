#! c:\perl\bin\perl.exe
#------------------------------------------------------------
# grab.pl
# Grab files (based on file list) from mounted image (via ImDisk,
#   SmartMount, etc.) or remote system mounted via F-Response
#
#
# copyright 2010 Quantum Analytics Research, LLC
#------------------------------------------------------------
use strict;
use Getopt::Long;
use Digest::MD5;
use File::Copy;

my $version = 20100607;

my %config = ();
Getopt::Long::Configure("prefix_pattern=(-|\/)");
GetOptions(\%config, qw(dir|d=s file|f=s pref|p user|u vista|v output|o=s help|?|h));

if ($config{help} || ! %config) {
	\_syntax();
	exit 1;
}

# Some error checking
die "You must provide a drive letter (E, F, G...).\n" unless ($config{dir});
my $drive = $config{dir}.":\\";
die "Drive ".$drive." not found.\n" unless (-e $drive && -d $drive);

die "You must provide an output directory path.\n" unless ($config{output});
die "Dir ".$config{output}." not found.\n" unless (-d $config{output} && -e $config{output});
my $out = $config{output};
$out .= "\\" unless ($out =~ m/\\$/);

my $filelist;
($config{file})?($filelist = $config{file}):($filelist = "files\.txt");

# Read in the file list; one file per line
my %files;
open(FH,"<",$filelist) || die "Could not open ".$filelist.": $!\n";
while(<FH>) {
# skip lines that start with # and lines that are blank	
	next if ($_ =~ m/^#/ || $_ =~ m/^\s+$/);
	chomp;
	$files{$drive.$_} = 1;
}
close(FH);

foreach my $f (keys %files) {
	if (-e $f && -f $f) {
		print "Processing $f...\n";
		my $hash1 = getHash($f);
		
		my @n = split(/\\/,$f);
		my $name = $n[scalar(@n) - 1];
		
		my $f2 = $out.$name;
		print "  Copying $f to $f2...\n";
		
		if (copy($f,$f2)) {
			my $hash2 = getHash($f2);
		
			if ($hash2 eq $hash1) {
				print "  --> Hashes match.\n";
				writeLog("$f successfully copied to $f2");
				writeLog("  $hash1 : $f");
				writeLog("  $hash2 : $f2");
				writeLog("  **Hashes match.");
			}
			else {
				writeLog("$f copied to $f2, hashes do not match.");
				print "  --> Hashes DO NOT match.\n";
			}
		}
		else {
#			print "Copy ".$f." to ".$f2." failed: $!\n";
			writeLog("Copy ".$f." to ".$f2." failed: $!");
		}
	}
	else {
#		print $f." not found.\n";
		writeLog($f." not found.");
	}
	
	writeLog("");
	if ($config{user}) {
		writeLog("------- Getting User Registry Hives -------");
		getUsers($drive);
	}
	writeLog("");
	if ($config{pref}) {
		writeLog("------- Getting User Registry Hives -------");
		getPrefetch($drive);
	}
}

print "Done.\n";

#------------------------------------------------------------
# functions
#------------------------------------------------------------
sub getUsers {
	my $d = shift;
	my $path;
	if ($config{vista}) {
		$path = $d."Users\\";
	}
	else {
		$path = $d."Documents and Settings\\";
	}
	my @users;
	if (-e $path && -d $path) {
		opendir(DIR, $path);
		@users = readdir(DIR);
		closedir(DIR);
	}
	else {
		writeLog($path." not found.");
		return;
	}
	
	foreach my $u (@users) {
		my $n = $path.$u."\\NTUSER\.DAT";
		if (-e $n && -f $n) {
			print "Processing $n...\n";
			my $hash1 = getHash($n);
		
			my $f2 = $out.$u."_NTUSER\.DAT";
			print "  Copying $n to $f2...\n";
		
			if (copy($n,$f2)) {
				my $hash2 = getHash($f2);
		
				if ($hash2 eq $hash1) {
					print "  --> Hashes match.\n";
					writeLog("$n successfully copied to $f2");
					writeLog("  $hash1 : $n");
					writeLog("  $hash2 : $f2");
					writeLog("  **Hashes match.");
				}
				else {
					writeLog("$n copied to $f2, hashes do not match.");
					print "  --> Hashes DO NOT match.\n";
				}
			}
			else {
#			print "Copy ".$f." to ".$f2." failed: $!\n";
				writeLog("Copy ".$n." to ".$f2." failed: $!");
			}
			writeLog("");
		}
	}
}

sub getPrefetch {
	my $d = shift;
	my $path = $d."Windows\\Prefetch\\";
	my @files;
	if (-e $path && -d $path) {
		opendir(DIR, $path);
		@files = readdir(DIR);
		closedir(DIR);
	}
	else {
		print $path." not found.\n";
		writeLog($path." not found.");
		return;
	}
	
	my $opath = $out."Prefetch\\";
	mkdir($opath,0777) unless (-e $opath && -d $opath);
	
	foreach my $pf (@files) {
		next unless ($pf =~ m/\.pf$/);
		my $f = $path.$pf;
		if (-e $f && -f $f) {
			print "Processing $f...\n";
			my $hash1 = getHash($f);
			my $f2 = $opath.$pf;
			print "  Copying $f to $f2...\n";
			if (copy($f,$f2)) {
				my $hash2 = getHash($f2);
				if ($hash2 eq $hash1) {
					print "  --> Hashes match.\n";
					writeLog("$f successfully copied to $f2");
					writeLog("  $hash1 : $f");
					writeLog("  $hash2 : $f2");
					writeLog("  **Hashes match.");
				}
				else {
					writeLog("$f copied to $f2, hashes do not match.");
					print "  --> Hashes DO NOT match.\n";
				}
			}
			else {
#			print "Copy ".$f." to ".$f2." failed: $!\n";
				writeLog("Copy ".$f." to ".$f2." failed: $!");
			}
		}
		else {
#		print $f." not found.\n";
			writeLog($f." not found.");
		}
		writeLog("");
	}
}

sub getHash {
	my $file = shift;
	my $digest;
	eval {
		open(FILE,"<",$file);
   	binmode(FILE);
   	$digest = Digest::MD5->new->addfile(*FILE)->hexdigest();
   	close(FH);
   	return $digest;
  };
}

sub writeLog {
	my $str = shift;
	my $o = $out."logfile\.txt";
	open(FH,">>",$o) || die "Could not open ".$o." to write: $!\n";
	print FH $str."\n";
	close(FH);
}


sub _syntax {
	print<< "EOT";
grab v.$version [-d drive] [-o output_dir][-f file][-puvh]
Grab files from mounted drive
 
  -d drive.....Drive letter (letter only; G)
  -f file......File list (default: files\.txt)
  -o output....Directory to save output
  -p ..........Get Prefetch files (creates Prefetch dir in
               output dir)
  -u ..........Get user Registry hives (prepends w/ user name) 
  -v...........Use Vista path (C:\\Users) for user profiles (default: 
               XP style path)             
  -h...........Help (print this information)

Ex: C:\\>grab -d G -f files\.txt -o E:\\test
    C:\\>grab -d G -f files\.txt -o E:\\test -u -p
  
copyright 2010 Quantum Analytics Research, LLC
EOT
}