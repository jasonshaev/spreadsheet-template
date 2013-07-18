package Spreadsheet::Template::Processor::Identity;
use Moose;
# ABSTRACT: render a template file with no processing at all

with 'Spreadsheet::Template::Processor';

=head1 SYNOPSIS

  my $template = Spreadsheet::Template->new(
      processor_class => 'Spreadsheet::Template::Processor::Identity',
  );

=head1 DESCRIPTION

This class implements L<Spreadsheet::Template::Processor>, and just passes
through the JSON data without modification.

=cut

sub process {
    my $self = shift;
    my ($contents, $vars) = @_;
    return $contents;
}

__PACKAGE__->meta->make_immutable;
no Moose;

=begin Pod::Coverage

  process

=end Pod::Coverage

=cut

1;
