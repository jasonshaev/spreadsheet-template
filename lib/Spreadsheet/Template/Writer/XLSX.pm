package Spreadsheet::Template::Writer::XLSX;
use Moose;
# ABSTRACT: generate XLSX files from templates

with 'Spreadsheet::Template::Writer::Excel';

=head1 SYNOPSIS

  my $template = Spreadsheet::Template->new(
      writer_class => 'Spreadsheet::Template::Writer::XLSX',
  );

=head1 DESCRIPTION

This class implements L<Spreadsheet::Template::Writer>, allowing you to
generate XLSX files.

=cut

sub excel_class { 'Excel::Writer::XLSX' }

__PACKAGE__->meta->make_immutable;
no Moose;

=begin Pod::Coverage

  excel_class

=end Pod::Coverage

=cut

1;
