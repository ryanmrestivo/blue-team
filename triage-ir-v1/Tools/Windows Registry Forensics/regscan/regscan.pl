#! c:\perl\bin\perl.exe
#---------------------------------------------------------
# regscan.pl
# Retrieves data from Windows Service Registry keys; LastWrite times,
#  ImagePath value (if avail.), Parameters\ServiceDll value (if avail),
#  and lists all entries sorted based on LastWrite times.
#
# usage: regscan.pl <system_name>
#
# Output:
#  LastWrite Time|ServiceName|ImagePath|ServiceDll
#  - values are "|" separated
#
# Copyright 2010 Quantum Analytics Research, LLC
#---------------------------------------------------------
use strict;
use Win32::TieRegistry(Delimiter=>"/");

my $server = shift || Win32::NodeName;
my $regkey = "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Services\\";
$regkey =~ s/\\/\//g;
$regkey = "//$server/".$regkey;

my %svcs;
my $remote;
eval {
	$remote = $Registry->Open($regkey, {Access=>0x20019});
};
die "Error occurred connecting to Registry: $@\n" if ($@);
	
# If connected to the key, dump a list of subkeys
my @subkeys = $remote->SubKeyNames();
foreach my $s (@subkeys) {
	my $str = $s;
		
	my %info = $remote->Information();
	my $lw = getTime(unpack("VV",$info{"LastWrite"}));
	
	eval {
		my $k = $remote->Open($s,{Access=>0x20019});
		$str .= "|".$k->GetValue("ImagePath");
	};
	$str .= "||" if ($@);
	
	eval {
		my $k = $remote->Open($s."\\Parameters",{Access=>0x20019});
		$str .= "|".$k->GetValue("ServiceDll");
	};
	$str .= "||" if ($@);
	
	my $type;
	eval {
		my $k = $remote->Open($s,{Access=>0x20019});
		$type = $k->GetValue("Type");
	};
	print "  ERROR: ".$@."\n" if ($@);
	
	push(@{$svcs{$lw}},$str) if ($type eq "0x00000010" || $type eq "0x00000020");
}
	
foreach my $t (reverse sort {$a <=> $b} keys %svcs) {
	foreach my $item (@{$svcs{$t}}) {
		print gmtime($t)."Z"."|".$item."\n";
	}
}

#-------------------------------------------------------------
# getTime()
# Translate FILETIME object (QWORD) to Unix time, to be passed
# to gmtime() or localtime()
#-------------------------------------------------------------
sub getTime() {
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