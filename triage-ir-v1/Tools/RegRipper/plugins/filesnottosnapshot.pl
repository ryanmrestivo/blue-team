#-----------------------------------------------------------
# filesnottosnapshot.pl
#   Access System hive file to get the contents of the FilesNotToSnapshot key
# 
# Change history
#   
#
# References
#   Troy Larson's Windows 7 presentation slide deck http://computer-forensics.sans.org/summit-archives/2010/files/12-larson-windows7-foreniscs.pdf
#   QCCIS white paper Reliably recovering evidential data from Volume Shadow Copies http://www.qccis.com/downloads/whitepapers/QCC%20VSS
#	http://msdn.microsoft.com/en-us/library/windows/desktop/bb891959(v=vs.85).aspx
# 
# copyright 2012 Corey Harrell
#-----------------------------------------------------------
package filesnottosnapshot;
use strict;

my %config = (hive          => "SYSTEM",
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              osmask        => 22,
              version       => 20120131);

sub getConfig{return %config}
sub getShortDescr {
	return "Get FilesNotToSnapshot key contents";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $hive = shift;
	::logMsg("Launching filesnottosnapshot v.".$VERSION);
    ::rptMsg("filesnottosnapshot v.".$VERSION); 
    ::rptMsg("(".getHive().") ".getShortDescr()."\n"); 
	
	my $reg = Parse::Win32Registry->new($hive);
	my $root_key = $reg->get_root_key;
# First thing to do is get the ControlSet00x marked current...this is
# going to be used over and over again in plugins that access the system
# file
	my $current;
	my $key_path = 'Select';
	my $key;
	if ($key = $root_key->get_subkey($key_path)) {
		$current = $key->get_value("Current")->get_data();
		my $ccs = "ControlSet00".$current;
		my $tz_path = $ccs."\\Control\\BackupRestore\\FilesNotToSnapshot";
		my $tz;
		if ($tz = $root_key->get_subkey($tz_path)) {
			::rptMsg("FilesNotToSnapshot key");
			::rptMsg($tz_path);
			::rptMsg("LastWrite Time ".gmtime($tz->get_timestamp())." (UTC)");
			::rptMsg("");
	
		my %cv;
		my @vals = $tz->get_list_of_values();;
		if (scalar(@vals) > 0) {
			foreach my $v (@vals) {
				my $name = $v->get_name();
				my $data = $v->get_data();
				my $len  = length($data);
				next if ($name eq "");
				push(@{$cv{$len}},$name." : ".$data);
			}
			foreach my $t (sort {$a <=> $b} keys %cv) {
				foreach my $item (@{$cv{$t}}) {
					::rptMsg("  $item");
				}
			}	
			::rptMsg("");
			::rptMsg("");
			::rptMsg("The listed directories/files are not backed up in Volume Shadow Copies");
		}
		else {
			::rptMsg($tz_path." has no values.");
			::logMsg($tz_path." has no values");
			}
		}
		else {
		::rptMsg($tz_path." not found.");
		::logMsg($tz_path." not found.");
		}
	}
	
}
	
1;	
	
	