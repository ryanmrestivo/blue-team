#-----------------------------------------------------------
# logon_xp_run
#   Get contents of Run key from Software hive
#
# Change history
#   20110830 [fpi] + banner, no change to the version number
#
# References
#   http://support.microsoft.com/kb/314488
#   
# Note: Needs testing to see if it applies beyond XP/XP-64
#
# copyright 2008 H. Carvey, keydet89@yahoo.com
#-----------------------------------------------------------
package logon_xp_run;
use strict;

my %config = (hive          => "NTUSER\.DAT",
              osmask        => 12,
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              version       => 20080328);

sub getConfig{return %config}

sub getShortDescr {
	return "Autostart - Get XP user logon Run key contents from NTUSER\.DAT hive";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $hive = shift;
	::logMsg("Launching user_xp_run v.".$VERSION);
    ::rptMsg("user_xp_run v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".getHive().") ".getShortDescr()."\n"); # 20110830 [fpi] + banner 
	my $reg = Parse::Win32Registry->new($hive);
	my $root_key = $reg->get_root_key;

	my $key_path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\policies\\Explorer\\Run";
	my $key;
	if ($key = $root_key->get_subkey($key_path)) {
		::rptMsg($key_path);
		::rptMsg("LastWrite Time ".gmtime($key->get_timestamp())." (UTC)");
		
		my %vals = getKeyValues($key);
		if (scalar(keys %vals) > 0) {
			foreach my $v (keys %vals) {
				::rptMsg("\t".$v." -> ".$vals{$v});
			}
		}
		else {
			::rptMsg($key_path." has no values.");
		}
		
#		my @sk = $key->get_list_of_subkeys();
#		if (scalar(@sk) > 0) {
#			foreach my $s (@sk) {
#				::rptMsg("");
#				::rptMsg($key_path."\\".$s->get_name());
#				::rptMsg("LastWrite Time ".gmtime($s->get_timestamp())." (UTC)");
#				my %vals = getKeyValues($s);
#				foreach my $v (keys %vals) {
#					::rptMsg("\t".$v." -> ".$vals{$v});
#				}
#			}
#		}
#		else {
#			::rptMsg("");
#			::rptMsg($key_path." has no subkeys.");
#		}
	}
	else {
		::rptMsg($key_path." not found.");
		::logMsg($key_path." not found.");
	}
	
}

sub getKeyValues {
	my $key = shift;
	my %vals;
	
	my @vk = $key->get_list_of_values();
	if (scalar(@vk) > 0) {
		foreach my $v (@vk) {
			next if ($v->get_name() eq "" && $v->get_data() eq "");
			$vals{$v->get_name()} = $v->get_data();
		}
	}
	else {
# do nothing	
	}
	return %vals;
}

1;