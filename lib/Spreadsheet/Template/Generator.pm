package Spreadsheet::Template::Generator;
use Moose;
# ABSTRACT: create new templates from existing spreadsheets

use Class::Load 'load_class';
use JSON;

has parser_class => (
    is      => 'ro',
    isa     => 'Str',
    default => 'Spreadsheet::Template::Generator::Parser::XLSX',
);

has parser_options => (
    is      => 'ro',
    isa     => 'HashRef',
    default => sub { {} },
);

has parser => (
    is   => 'ro',
    does => 'Spreadsheet::Template::Generator::Parser',
    lazy => 1,
    default => sub {
        my $self = shift;
        my $class = $self->parser_class;
        load_class($class);
        return $class->new(
            %{ $self->parser_options }
        );
    },
);

sub generate {
    my $self = shift;
    my ($filename) = @_;
    my $data = $self->parser->parse($filename);
    return JSON->new->pretty->canonical->encode($data);
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
