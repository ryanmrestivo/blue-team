#! c:\perl\bin\perl.exe
#-----------------------------------------------------------
# maclookup.pl
# code to translate WAP MAC addresses acquired from within Windows
# Registry entries into map locations
#
# based on: http://spl0it.org/files/bssid-location.pl
# *requires Crypt::SSLeay for the HTTPS communications
# 
# Once lat/longs are acquired, submit to Google Maps as follows:
# http://maps.google.com/maps?q=$lat,+$long+%28$str%29&iwloc=A&hl=en
# 
# Future work:
#   - In order to do multiple pushpins on a single map, would need 
#     to create a KML or KMZ file
#
# copyright 2010 Quantum Analytics Research, LLC
#
# This script is provided 'as is' with no guarantees or warantees as to
# its functionality.  Use at your own risk.
#-----------------------------------------------------------
use strict;
use LWP::UserAgent; # also need Crypt::SSLeay for HTTPS comms
use XML::Simple;
#-----------------------------------------------------------
# http://search.cpan.org/~bdfoy/Net-MAC-Vendor-1.18/lib/Vendor.pm
#
# To install, copy Vendor.pm into the site\lib\Net\MAC\ folder in your
# Perl distro (may have to create "MAC").
#-----------------------------------------------------------
use Net::MAC::Vendor;
use Getopt::Long;
my %config = ();
Getopt::Long::Configure("prefix_pattern=(-|\/)");
GetOptions(\%config, qw(file|f=s ssid|s=s wap|w=s help|?|h));

if ($config{help} || ! %config) {
	_syntax();
	exit 1;
}

my %macs;

#---------------------------------------------------------
# If WAP MACs and SSIDs are placed in a file, put MAC and SSID
# on a single line, separated by a semi-colon.
#---------------------------------------------------------
if ($config{file}) {
	die $config{file}." not found.\n" unless (-e $config{file} && -f $config{file});
	open(FH,"<",$config{file}) || die "Could not open ".$config{file}."\n";
	while(<FH>) {
		chomp;
		my ($wap,$ssid) = split(/;/,$_,2);
		$ssid = "Unknown" if ($ssid eq "");
		$macs{$wap} = $ssid;
	}
	close(FH);
}

if ($config{wap}) {
	($config{ssid}) ? ($macs{$config{wap}} = $config{ssid}) : ($macs{$config{wap}} =  "Unknown");
}

foreach my $mac (keys %macs) {
	my $ssid = $macs{$mac};
# if MAC address contains dashes, convert to colons
	if (length($mac) == 17) {
		if (index($mac,"-") != -1) {
			$mac =~ s/\-/\:/g;
		}
	}
	elsif (length($mac) == 12) {
		
	}
	else {}
	
	eval {
		print "OUI lookup for $mac...\n";
		my $macs = Net::MAC::Vendor::lookup($mac);
		foreach (@$macs) {
			print "  ".$_."\n";
		}
	};
	print "Error looking up OUI: $@\n" if ($@);
	print "\n";
	
# Remove colons from MAC address
	$mac =~ s/\://g;
	die "MAC address is wrong size.\n" if (length($mac) != 12);
 
	my $url = 'https://api.skyhookwireless.com/wps2/location';
	my $ua = LWP::UserAgent->new;
	$ua->timeout(10);
	$ua->env_proxy;

	my $str = "<?xml version='1.0'?>  
 <LocationRQ xmlns='http://skyhookwireless.com/wps/2005' version='2.6' street-address-lookup='full'>  
   <authentication version='2.0'>  
     <simple>  
       <username>beta</username>  
       <realm>js.loki.com</realm>  
     </simple>  
   </authentication>  
   <access-point>  
     <mac>$mac</mac>  
     <signal-strength>-50</signal-strength>  
   </access-point>  
 </LocationRQ>";

	my ($lat,$long);
	eval {
		my $resp = $ua->post($url,'Content-Type' => 'text/xml',Content => $str);
		($lat,$long) = getLatLong($resp);
	};
	print "Error looking up lat/long: $@\n" if ($@);
#	print "Latitude  = ".$lat."\n";
#	print "Longitude = ".$long."\n";
	if ($lat ne "" && $long ne "") {
		my $mapurl = "http://maps.google.com/maps?q=".$lat.",+".$long."+%28".$ssid."%29&iwloc=A&hl=en";
		print "Google Map URL (paste into browser):\n";
		print $mapurl."\n";
	}
	else {
		print "Apparently, no lat/longs are available for ".$mac."\n";
	}
	print "\n";
}

sub getLatLong {
	my ($response) = shift;
  my $xml = $response->content;
# print $xml."\n";
	my $config = XMLin($xml);	
#	print "Latitude  : ".$config->{location}->{latitude}."\n";
#	print "Longitude : ".$config->{location}->{longitude}."\n";
	return ($config->{location}->{latitude},$config->{location}->{longitude});
}

#---------------------------------------------------------
# _syntax()
# 
#---------------------------------------------------------
sub _syntax {
print<< "EOT";
maclookup [options]
Translate WAP MAC addresses acquired from within Windows Registry entries 
into map locations

  -f file........read MAC addresses (and SSIDs) from a file
  -s SSID........WAP SSID          
  -w MAC addr....WAP MAC address (octets separated by dashes or colons)
  -h ............Help (print this information)
  
Ex: C:\\>maclookup -f <path_to_file>
    C:\\>maclookup -w 00-19-07-5B-36-92 -s tmobile

copyright 2010 Quantum Analytics Research, LLC
EOT
}