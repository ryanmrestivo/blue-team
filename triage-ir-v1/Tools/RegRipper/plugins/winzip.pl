#-----------------------------------------------------------
# winzip.pl
#
# Change History
#   20110830 [fpi] + banner, no change to the version number
#
# References
#
# copyright 2008 H. Carvey, keydet89@yahoo.com
#-----------------------------------------------------------
package winzip;
use strict;

my %config = (hive          => "NTUSER\.DAT",
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              osmask        => 22,
              version       => 20080325);

sub getConfig{return %config}
sub getShortDescr {
	return "Get WinZip extract and filemenu values";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $hive = shift;
	::logMsg("Launching winzip v.".$VERSION);
    ::rptMsg("winzip v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".getHive().") ".getShortDescr()."\n"); # 20110830 [fpi] + banner

	my $reg = Parse::Win32Registry->new($hive);
	my $root_key = $reg->get_root_key;
	my $key_path = "Software\\Nico Mak Computing\\WinZip";
	my $key;
	if ($key = $root_key->get_subkey($key_path)) {
		::rptMsg("WinZip");
		::rptMsg($key_path);
		::rptMsg("");
		my @subkeys = $key->get_list_of_subkeys();
		my %sk;
		foreach my $s (@subkeys) {
			$sk{$s->get_name()} = $s;
		}
		
		if (exists $sk{'extract'}) {
			my $tag = "extract";
			::rptMsg($key_path."\\extract  [".gmtime($sk{'extract'}->get_timestamp)."]");
			my @vals = $sk{'extract'}->get_list_of_values();
			my %ext;
			foreach my $v (@vals) {
				my $name = $v->get_name();
				my $num = $name;
				$num =~ s/^$tag//;
				$ext{$num} = $v->get_data();
			}
			foreach my $e (sort {$a <=> $b} keys %ext) {
				::rptMsg("  extract".$e." -> ".$ext{$e});
			}
			::rptMsg("");
		}
		else {
			::rptMsg("extract key not found.");
		}
		
		if (exists $sk{'filemenu'}) {
			my $tag = "filemenu";
			::rptMsg($key_path."\\filemenu  [".gmtime($sk{'extract'}->get_timestamp)."]");
			my @vals = $sk{'filemenu'}->get_list_of_values();
			my %ext;
			foreach my $v (@vals) {
				my $name = $v->get_name();
				my $num = $name;
				$num =~ s/^$tag//;
				$ext{$num} = $v->get_data();
			}
			foreach my $e (sort {$a <=> $b} keys %ext) {
				::rptMsg("  filemenu".$e." -> ".$ext{$e});
			}
		}
		else {
			::rptMsg("filemenu key not found.");
		}
	}
	else {
		::rptMsg($key_path." not found.");
		::logMsg($key_path." not found.");
	}
}
1;