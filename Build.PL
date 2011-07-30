use 5.006;
use strict;
use warnings;
use Module::Build;

my $builder = Module::Build->new(
    module_name         => 'C3000',
    license             => 'perl',
    dist_author         => q{Andy Xiao <xyf.xiao@gmail.com>},
    dist_version_from   => 'lib/C3000.pm',
    build_requires => {
        'Test::More' => 0,
    },
    requires => {
        'perl' => 5.006,
    },
    add_to_cleanup      => [ 'C3000-*' ],
    create_makefile_pl => 'traditional',
);

$builder->create_build_script();
