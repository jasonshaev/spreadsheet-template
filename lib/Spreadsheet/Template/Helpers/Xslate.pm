package Spreadsheet::Template::Helpers::Xslate;
use strict;
use warnings;
use Data::Dumper;

use JSON;

my $JSON = JSON->new;

use Sub::Exporter 'build_exporter';

my $import = build_exporter({
    exports => [
        map { $_ => \&_curry_package } qw(format c merge true false)
    ],
    groups => {
        default => [qw(format c merge true false)],
    },
});

my %formats;

sub import {
    my $caller = caller;
    $formats{$caller} = {};
    goto $import;
}

sub format {
    my ($package, $name, $format) = @_;
    $formats{$package}{$name} = $format;
    return '';
}

sub merge {
    my ($package, $range, $contents, $format) = @_;

    return $JSON->encode({
        range    => "$range",
        contents => "$contents",
        format   => _formats($package, $format)
    });
}

sub c {
    my ($package, $contents, $format, $type, %args) = @_;

    $type = 'string' unless defined $type;

    return $JSON->encode({
        contents => "$contents",
        format   => _formats($package, $format),
        type     => $type,
        (defined $args{formula}
            ? (formula => $args{formula})
            : ()),
    });
}

sub true  { JSON::true  }
sub false { JSON::false }

sub _formats {
    my ($package, $format) = @_;

    return $format if ref($format);

    return $formats{$package}{$format};
}

sub _curry_package {
    my ($package, $name) = @_;
    return sub { $package->$name(@_) };
}

=begin Pod::Coverage

 format
 c
 true
 false

=end Pod::Coverage

=cut

1;
