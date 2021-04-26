#-----------------------------------------------------------
# comdlg32.pl
#   Plugin for Registry Ripper 
#
# Change history
#   20080324 [hca] % created
#   20100402 [hca] * updated IAW Chad Tilbury's post to SANS Forensic Blog
#   20100409 [hca] * updated to include Vista and above
#   20110830 [fpi] + banner, no change to the version number
#   20110901 [fpi] * eliminated old "comdlg32.pl" and renamed this plugin
#                    "comdlg32a.pl" to "comdlg32.pl". Updated version number
#                    to be able to compare and track down changes
#
# References
#   Win2000 - http://support.microsoft.com/kb/319958
#   XP - http://support.microsoft.com/kb/322948/EN-US/
#		
# copyright 20100402 Quantum Analytics Research, LLC
#-----------------------------------------------------------
package comdlg32;
use strict;

my %config = (hive          => "NTUSER\.DAT",
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              osmask        => 22,
              version       => 20110901);

sub getConfig{return %config}
sub getShortDescr {
	return "Gets contents of user's ComDlg32 key";
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $ntuser = shift;
	::logMsg("Launching comdlg32 v.".$VERSION);
    ::rptMsg("comdlg32 v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".$config{hive}.") ".getShortDescr()."\n"); # 20110830 [fpi] + banner 
	my $reg = Parse::Win32Registry->new($ntuser);
	my $root_key = $reg->get_root_key;
	::rptMsg("comdlg32 v.".$VERSION);
	
# LastVistedMRU	
	my $key_path = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ComDlg32";
	my $key;
	my @vals;
	if ($key = $root_key->get_subkey($key_path)) {
		::rptMsg($key_path);
		::rptMsg("LastWrite Time ".gmtime($key->get_timestamp())." (UTC)");
		
		my @subkeys = $key->get_list_of_subkeys();
		
		if (scalar @subkeys > 0) {
			foreach my $s (@subkeys) {
				parseLastVisitedMRU($s) if ($s->get_name() eq "LastVisitedMRU");
				parseOpenSaveMRU($s) if ($s->get_name() eq "OpenSaveMRU");
			}
		}
		else {
			::rptMsg($key_path." has no subkeys.");
		}
	}
}	

sub parseLastVisitedMRU {
	my $key = shift;
	my %lvmru;
	my @mrulist;
	my @vals = $key->get_list_of_values();
	
	if (scalar(@vals) > 0) {
# First, read in all of the values and the data
		foreach my $v (@vals) {
			$lvmru{$v->get_name()} = $v->get_data();
		}
# Then, remove the MRUList value
		if (exists $lvmru{MRUList}) {
			::rptMsg("  MRUList = ".$lvmru{MRUList});
			@mrulist = split(//,$lvmru{MRUList});
			delete($lvmru{MRUList});
			foreach my $m (@mrulist) {
				my ($file,$dir) = split(/\00\00/,$lvmru{$m},2);
				$file =~ s/\00//g;
				$dir  =~ s/\00//g;
				::rptMsg("  ".$m." -> EXE: ".$file);
				::rptMsg("    -> Last Dir: ".$dir);
			}
		}
		else {
			::rptMsg("LastVisitedMRU key does not have an MRUList value.");
		}				
	}
	else {
		::rptMsg("LastVisitedMRU key has no values.");
	}	
	::rptMsg("");
}

sub parseOpenSaveMRU {
	my $key = shift;
	
	parseOpenSaveValues($key);
	::rptMsg("");
# Now, let's get the subkeys
	my @sk = $key->get_list_of_subkeys();
	if (scalar(@sk) > 0) {
		foreach my $s (@sk) {
			parseOpenSaveValues($s);
			::rptMsg("");
		}
	}
	else {
		::rptMsg("OpenSaveMRU key has no subkeys.");
	}	
	::rptMsg("");
}

sub parseOpenSaveValues {
	my $key = shift;
	::rptMsg("OpenSaveMRU\\".$key->get_name());
	::rptMsg("LastWrite Time: ".gmtime($key->get_timestamp())." Z");
	my %osmru;
	my @vals = $key->get_list_of_values();
	if (scalar(@vals) > 0) {
		map{$osmru{$_->get_name()} = $_->get_data()}(@vals);
		if (exists $osmru{MRUList}) {
			::rptMsg("  MRUList = ".$osmru{MRUList});
			my @mrulist = split(//,$osmru{MRUList});
			delete($osmru{MRUList});
			foreach my $m (@mrulist) {
				::rptMsg("  ".$m." -> ".$osmru{$m});
			}
		}
		else {
			::rptMsg($key->get_name()." does not have an MRUList value.");
		}	
	}
	else {
		::rptMsg($key->get_name()." has no values.");
	}	
}

sub parseCIDSizeMRU {
	my $key = shift;
	my %lvmru;
	my @mrulist;
	my @vals = $key->get_list_of_values();
		
	if (scalar(@vals) > 0) {
# First, read in all of the values and the data
		foreach my $v (@vals) {
			$lvmru{$v->get_name()} = $v->get_data();
		}
# Then, remove the MRUList value
		if (exists $lvmru{MRUListEx}) {
			delete($lvmru{MRUListEx});
			foreach my $m (keys %lvmru) {
				my $file = parseStr($lvmru{$m});
				my $str = sprintf "%-4s ".$file,$m;
				::rptMsg("  ".$str);
			}
		}
		else {
			::rptMsg("CIDSizeMRU does not have an MRUList value.");
		}				
	}
	else {
		::rptMsg("CIDSizeMRU has no values.");
	}
}

sub parseLastVisitedPidlMRU {
	my $key = shift;
	my %lvmru;
	my @mrulist;
	my @vals = $key->get_list_of_values();
	
	if (scalar(@vals) > 0) {
# First, read in all of the values and the data
		foreach my $v (@vals) {
			$lvmru{$v->get_name()} = $v->get_data();
		}
# Then, remove the MRUList value
		if (exists $lvmru{MRUListEx}) {
			delete($lvmru{MRUListEx});
			foreach my $m (keys %lvmru) {
				my $file = parseStr($lvmru{$m});
				my $str = sprintf "%-4s ".$file,$m;
				::rptMsg("  ".$str);
			}
		}
		else {
			::rptMsg("LastVisitedPidlMRU key does not have an MRUList value.");
		}				
	}
	else {
		::rptMsg("LastVisitedPidlMRU key has no values.");
	}	
}

sub parseStr {
	my $data = $_[0];
	my $temp;
	my $tag = 1;
	my $ofs = 0;
	
	while ($tag) {
		my $t = substr($data,$ofs,2);
		if (unpack("v",$t) == 0x00) {
			$tag = 0;
		}
		else {
			$temp .= $t;
			$ofs += 2;
		}
	}
	$temp =~ s/\00//g;
	return $temp;
}

1;		