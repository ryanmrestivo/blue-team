#-----------------------------------------------------------
# userassist2.pl
#   Gets the UserAssist values, NTUSER.DAT edition
#
# Change history
#   20100308 [qar] % created, based on original userassist.pl plugin
#   20100322 [qar] + added CLSID list reference
#   20110830 [fpi] + banner, no change to the version number
#  
# References
#   Control Panel Applets - http://support.microsoft.com/kb/313808
#   CLSIDs - http://www.autohotkey.com/docs/misc/CLSID-List.htm
# 
# copyright 2010 Quantum Analytics Research, LLC
#-----------------------------------------------------------
package userassist2;
use strict;

my %config = (hive          => "NTUSER\.DAT",
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              osmask        => 22,
              version       => 20100308);

sub getConfig{return %config}
sub getShortDescr {
	return "Displays contents of UserAssist subkeys";	
}
sub getDescr{}
sub getRefs {"Description of Control Panel Files in XP" => "http://support.microsoft.com/kb/313808"}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $ntuser = shift;
	::logMsg("Launching userassist2 v.".$VERSION);
    ::rptMsg("userassist2 v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".getHive().") ".getShortDescr()."\n"); # 20110830 [fpi] + banner

	my $reg = Parse::Win32Registry->new($ntuser);
	my $root_key = $reg->get_root_key;
	
	my $key_path = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\UserAssist";              
	my $key;
	
	if ($key = $root_key->get_subkey($key_path)) {
		::rptMsg("UserAssist");
		::rptMsg($key_path);
		::rptMsg("LastWrite Time ".gmtime($key->get_timestamp())." (UTC)");
		::rptMsg("");
		my @subkeys = $key->get_list_of_subkeys();
		if (scalar(@subkeys) > 0) {
			foreach my $s (@subkeys) {
				::rptMsg($s->get_name());
				processKey($s);
				::rptMsg("");
			}
		}
		else {
			::rptMsg($key_path." has no subkeys.");
		}
	}
	else {
		::rptMsg($key_path." not found.");
	}
}

sub processKey {
	my $ua = shift;
	
	my $key = $ua->get_subkey("Count");

	my %ua;
	my $hrzr = "HRZR";
	
	my @vals = $key->get_list_of_values();
		if (scalar(@vals) > 0) {
			foreach my $v (@vals) {
				my $value_name = $v->get_name();
				my $data = $v->get_data();

# Windows XP/2003/Vista/2008
				if (length($data) == 16) {
					my ($session,$count,$val1,$val2) = unpack("V*",$data);
				 	if ($val2 != 0) {
						my $time_value = ::getTime($val1,$val2);
						if ($value_name =~ m/^$hrzr/) { 
							$value_name =~ tr/N-ZA-Mn-za-m/A-Za-z/;
						}
						$count -= 5 if ($count > 5);
						push(@{$ua{$time_value}},$value_name." (".$count.")");
					}
				}
# Windows 7				
				elsif (length($data) == 72) { 
					$value_name =~ tr/N-ZA-Mn-za-m/A-Za-z/;
					
					if (unpack("V",substr($data,0,4)) == 0) {
						
						my $count = unpack("V",substr($data,4,4));
						my @t = unpack("VV",substr($data,60,8));
						next if ($t[0] == 0 && $t[1] == 0);
						my $time_val = ::getTime($t[0],$t[1]);	
						push(@{$ua{$time_val}},$value_name." (".$count.")");
					}
				}
				else {
# Nothing else to do
				}
			}
			foreach my $t (reverse sort {$a <=> $b} keys %ua) {
				::rptMsg(gmtime($t)." Z");
				foreach my $i (@{$ua{$t}}) {
					::rptMsg("  ".$i);
				}
			}
		}
	
	
	
}







1;