#-----------------------------------------------------------
# networkcards.pl
# 
# Change history
#   20110830 [fpi] + banner, no change to the version number
#   20111024 [fpi] + added print of 'ServiceName' to correlate
#                    with other plugins
#
# References
# 
# copyright 2008 H. Carvey
#-----------------------------------------------------------
package networkcards;
use strict;

my %config = (hive          => "Software",
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              osmask        => 22,
              version       => 20080325);

sub getConfig{return %config}
sub getShortDescr {
	return "Get NetworkCards";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $hive = shift;
	::logMsg("Launching networkcards v.".$VERSION);
    ::rptMsg("networkcards v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".getHive().") ".getShortDescr()."\n"); # 20110830 [fpi] + banner

	my $reg = Parse::Win32Registry->new($hive);
	my $root_key = $reg->get_root_key;
	my $key_path = "Microsoft\\Windows NT\\CurrentVersion\\NetworkCards";
	my $key;
	if ($key = $root_key->get_subkey($key_path)) {
		::rptMsg("NetworkCards");
		::rptMsg($key_path);
		::rptMsg("");
		my @subkeys = $key->get_list_of_subkeys();
		if (scalar(@subkeys) > 0) {
			my %nc;
			foreach my $s (@subkeys) {
				my $service = $s->get_value("ServiceName")->get_data();
				$nc{$service}{descr} = $s->get_value("Description")->get_data();
				$nc{$service}{lastwrite} = $s->get_timestamp();
			}
			
			foreach my $n (keys %nc) {
				# WAS: ::rptMsg($nc{$n}{descr}."  [".gmtime($nc{$n}{lastwrite})."]");
				::rptMsg("$n ".$nc{$n}{descr}." [".gmtime($nc{$n}{lastwrite})."]"); # 20111024 [fpi]
			}
		}
		else {
			::rptMsg($key_path." has no subkeys.");
			::logMsg($key_path." has no subkeys.");
		}
	}
	else {
		::rptMsg($key_path." not found.");
		::logMsg($key_path." not found.");
	}
}
1;