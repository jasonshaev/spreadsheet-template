package Spreadsheet::Template::Generator;
use Moose;
# ABSTRACT: create new templates from existing spreadsheets

use Class::Load 'load_class';
use JSON;

has parser_options => (
    is      => 'ro',
    isa     => 'HashRef',
    default => sub { {} },
);

sub generate {
    my $self = shift;
    my ($filename) = @_;
    (my $ext = $filename) =~ s/.*\.//;
    my $class = $self->parser_classes->{$ext};
    load_class($class);
    my $parser = $class->new(
        filename => $filename,
        %{ $self->parser_options }
    );
    my $data = $parser->parse;
    return JSON->new->pretty->canonical->encode($data);
}

sub parser_classes {
    +{
        'xls'  => 'Spreadsheet::Template::Generator::Parser::XLS',
        'xlsx' => 'Spreadsheet::Template::Generator::Parser::XLSX',
    }
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
