#-----------------------------------------------------------
# installedcomp.pl
#   Get info about Installed Components
#
# Change history
#   20100115 [qar] % created
#   20100116 [qar] * updated for slightly better coverage
#   20110830 [fpi] + banner, no change to the version number
#
# References
#   
# Notes: Look for out of place entries, particularly those 
#        that point to the Recycle Bin or a temp directory
#
# copyright 2010 Quantum Analytics Research, LLC
#-----------------------------------------------------------
package installedcomp;
use strict;

my %config = (hive          => "Software",
              osmask        => 22,
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              version       => 20100116);

sub getConfig{return %config}

sub getShortDescr {
	return "Get info about Installed Components/StubPath";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

my %comp;

sub pluginmain {
	my $class = shift;
	my $hive = shift;
	::logMsg("Launching installedcomp v.".$VERSION);
    ::rptMsg("installedcomp v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".getHive().") ".getShortDescr()."\n"); # 20110830 [fpi] + banner
	my $reg = Parse::Win32Registry->new($hive);
	my $root_key = $reg->get_root_key;

	my $key_path = "Microsoft\\Active Setup\\Installed Components";
	my $key;
	if ($key = $root_key->get_subkey($key_path)) {
		::rptMsg("");
		::rptMsg($key_path);
		::rptMsg("LastWrite Time ".gmtime($key->get_timestamp())." (UTC)");
		::rptMsg("");
		
		my @subkeys = $key->get_list_of_subkeys();
		if (scalar(@subkeys) > 0) {
			foreach my $s (@subkeys) {
				my $lastwrite = $s->get_timestamp();
							
				my $str;
				eval {
					$str = $s->get_value("ComponentID")->get_data();
				};
				
				eval {
					my $ver = $s->get_value("Version")->get_data();
					$str .= " v.".$ver if ($ver && $s->get_value("Version")->get_type() == 1);
				};
				
				eval {
					my $stub = $s->get_value("StubPath")->get_data();
					$str .= "; ".$stub if ($stub ne "");
				};

# If the $str scalar is empty at this point, that means that for
# some reason, we haven't been able to populate the information
# we're looking for; in this case, we'll go looking for some info
# in a different area of the hive; the BHO.pl plugin does this, as
# well.  I'd rather that the plugin look for the Classes info than
# leave a blank entry in the output.
				if ($str eq "") {
					my $name = $s->get_name();
					my $class_path = "Classes\\CLSID\\".$name;
					my $proc;
					if ($proc = $root_key->get_subkey($class_path)) {
# Try these two eval{} statements because I've seen the different 
# spellings for InProcServer32/InprocServer32 in sequential keys
						eval {
							$str = $proc->get_subkey("InprocServer32")->get_value("")->get_data();
						};
						
						eval {
							$str = $proc->get_subkey("InProcServer32")->get_value("")->get_data();
						};
					}
					else {
						$str = $name." class not found.";
					}
				}
				
				push(@{$comp{$lastwrite}},$str);
			}

			foreach my $t (reverse sort {$a <=> $b} keys %comp) {
				::rptMsg(gmtime($t)." (UTC)");
				foreach my $item (@{$comp{$t}}) {
					::rptMsg("  ".$item);
				}
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
1;