#-----------------------------------------------------------
# regtime_tln.pl
#   Plugin for Registry Ripper; traverses through a Registry 
#   hive file, pulling out keys and their LastWrite times, and
#   then listing them in order, sorted by the most recent time
#   first - works with any Registry hive file.
#
# Change history
#   20110830 [fpi] + banner, no change to the version number
# 
# References
#
# copyright 2008 H. Carvey
#-----------------------------------------------------------
package regtime_tln;
use strict;

my %config = (hive          => "All",
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              osmask        => 22,
              version       => 20080324);

sub getConfig{return %config}
sub getShortDescr {
	return "Dumps entire hive - all keys sorted by LastWrite time";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

my %regkeys;

sub pluginmain {
	my $class = shift;
	my $file = shift;
    
	::logMsg("Launching regtime_tln v.".$VERSION);
    ::rptMsg("regtime_tln v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".getHive().") ".getShortDescr()."\n"); # 20110830 [fpi] + banner

	my $reg = Parse::Win32Registry->new($file);
	my $root_key = $reg->get_root_key;
	
	traverse($root_key);

	foreach my $t (reverse sort {$a <=> $b} keys %regkeys) {
		foreach my $item (@{$regkeys{$t}}) {
			#::rptMsg(gmtime($t)."Z \t".$item);
			::rptMsg($t."|".$file."|".$item."|0|0");
		}
	}
}

sub traverse {
	my $key = shift;
  my $ts = $key->get_timestamp();
  my $name = $key->as_string();
  $name =~ s/\$\$\$PROTO\.HIV//;
  $name = (split(/\[/,$name))[0];
  push(@{$regkeys{$ts}},$name);  
	foreach my $subkey ($key->get_list_of_subkeys()) {
		traverse($subkey);
  }
}

1;