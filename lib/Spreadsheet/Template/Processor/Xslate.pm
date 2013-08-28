package Spreadsheet::Template::Processor::Xslate;
use Moose;
# ABSTRACT: preprocess templates with Xslate

use Text::Xslate;

with 'Spreadsheet::Template::Processor';

=head1 SYNOPSIS

  my $template = Spreadsheet::Template->new(
      processor_class   => 'Spreadsheet::Template::Processor::Xslate',
      processor_options => {
          syntax => 'TTerse'
      },
  );

=head1 DESCRIPTION

This class implements L<Spreadsheet::Template::Processor> to run the template
data through L<Text::Xslate>. In addition to allowing you to use the provided
variables, it also provides some convenience macros to use when writing your
templates:

=over 4

=item format($name, $options)

Declares a named format, which can be used with the C<c> helper. C<$name> is
the name to use for the format, and C<$options> is a hashref to use as the
value for the C<format> entry in the cell.

=item c($contents, $format, $type, %args)

Returns the representation of a cell. C<$contents> is the cell contents,
C<$format> is either the name of a format declared with the C<format> helper,
or a hashref of format options, C<$type> is either C<"string">, C<"number">, or
C<"date_time">, and C<%args> contains any other parameters (such as C<formula>,
for instance) to declare for the cell. C<$type> is optional, and if not passed,
defaults to C<"string">.

=item merge($range, $content, $format, $type, %args)

Returns representation of a range of cells to be merged. C<$content>,
C<$format>, C<$type>, and C<%args> are identical to the parameters listed above
for the C<c> helper, and C<$range> describes the range of cells to be merged.
The range can be specified either by an array of two arrays corresponding to
the row and column indices of the top left and bottom right cell, or by an
Excel-style range (like C<A1:C3>).

=item true

Returns C<JSON::true>.

=item false

Returns C<JSON::false>.

=back

=cut

=attr syntax

Which Xslate syntax engine to use. Defaults to C<Metakolon>.

=cut

has syntax => (
    is      => 'ro',
    isa     => 'Str',
    default => 'Metakolon',
);

has xslate => (
    is      => 'ro',
    isa     => 'Text::Xslate',
    lazy    => 1,
    default => sub {
        my $self = shift;
        return Text::Xslate->new(
            type   => 'text',
            syntax => $self->syntax,
            module => ['Spreadsheet::Template::Helpers::Xslate'],
        );
    },
);

sub process {
    my $self = shift;
    my ($contents, $vars) = @_;
    return $self->xslate->render_string($contents, $vars);
}

__PACKAGE__->meta->make_immutable;
no Moose;

=begin Pod::Coverage

  process

=end Pod::Coverage

=cut

1;
