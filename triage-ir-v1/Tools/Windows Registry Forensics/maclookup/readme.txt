This is the readme file for maclookup.pl.  

Please be sure to read this file before attempting to use maclookup.pl.

While preparing this code for distribution on the DVD that accompanies the 
book, I found that I had a great deal of difficulty "compiling" the code 
with Perl2Exe, due to module dependencies.  As such, I opted to leave the
code as a Perl script.

This script relies on the following modules:
- LWP::UserAgent (you also need Crypt::SSLeay for HTTPS communications)
- XML::Simple
- Net::MAC::Vendor (get it from: 
http://search.cpan.org/~bdfoy/Net-MAC-Vendor-1.18/lib/Vendor.pm

To install, copy Vendor.pm into the site\lib\Net\MAC\ folder in your Perl 
distribution; you may have to create the "MAC" directory)

Note: this code makes use of the SkyHook Wireless database.  If you enter
a WAP MAC address, and don't get anything back, it may be that the address
is not in the SkyHook database.

This Perl code is provided AS-IS, with no warantees or guarantees as to its
functionality.  Please feel free to view, use, and modify the code as you like,
but please provide proper credit and attribution.