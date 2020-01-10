#!c:/Perl/perl.exe -w
use File::Copy;

# Criar layout de data
my($dd,$mm,$yy,$day,$hh,$nn) = (localtime)[3,4,5,6,2,1];
my $today =  join '', map sprintf("%02d", $_),($yy%100,$mm+1,$dd,);
my $hr = join '', map sprintf("%02d", $_),($hh,$nn);
my $data = $today.'_'.$hr;
my $ARQUIVOS = 'arquivos.txt';
my $Caminho = "BKP";


open(my $file, q{<}, $ARQUIVOS) or die "Can't open file $ARQUIVOS: $!\n";
foreach my $arquivo ( <$file> ) {
   #Remove o último caractere apenas se for igual a $/ "Separador de regsitro" 
   chomp($arquivo);

   # Verifica se o arquivo existe  s
   if (-e $arquivo) 
   {
	   mkdir $Caminho;
	   copy($arquivo,$Caminho."\\".$data."_".$arquivo) or die "Copy failed: $!";
   }
}