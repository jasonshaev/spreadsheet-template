package Spreadsheet::Template::Generator;
use Moose;

use Class::Load 'load_class';
use JSON;

sub generate {
    my $self = shift;
    my ($filename) = @_;
    (my $ext = $filename) =~ s/.*\.//;
    my $class = $self->parser_classes->{$ext};
    load_class($class);
    my $parser = $class->new;
    my $data = $parser->parse($filename);
    return JSON->new->pretty->encode($data);
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
