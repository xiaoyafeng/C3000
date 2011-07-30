#!perl -T
use C3000;
use Test::More;

my $hl = C3000->new();

my $meter = '%';
my $meter_var  = '+A';
my $return_fields = ['ADAS_DEVICE', 'ADAS_VARIABLE', 'ADAS_DEVICE_NAME','ADAS_VARIABLE_NAME'];
my $criteria = [['ADAS_DEVICE_NAME',  $meter], ['ADAS_VARIABLE_NAME', $meter_var]];
my $rs_device;
my $rs_LP;
ok($rs_device = $hl->search_device($return_fields, $criteria), "search_device test");

$return_fields = [ 'ADAS_TIME_GMT', 'ADAS_VAL_RAW', 'ADAS_USER_STATUS'];
my $device_id = $rs_device->Fields('ADAS_DEVICE')->{Value};
my $variable_id = $rs_device->Fields('ADAS_VARIABLE')->{Value};
my $date_from  = $hl->convert_VT_DATE('yesterday');
my $date_to    = $hl->convert_VT_DATE('today');
$criteria = [['ADAS_DEVICE', $device_id], ['ADAS_VARIABLE', $variable_id], ['ADAS_TIME_GMT', $date_from, '>'], ['ADAS_TIME_GMT', $date_to, '<=']];

ok($rs_LP = $hl->get_LP($return_fields, $criteria), "get load profile");

ok(defined($hl->accu_LP('ADAS_VAL_RAW', 'º£%', '+A', 'yesterday', 'today')), "get accumulation");
ok(defined($hl->get_single_LP('ADAS_VAL_RAW', 'º£%', '+A', 'yesterday')), "get single value");






done_testing;
