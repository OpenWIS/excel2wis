#!/usr/bin/perl
use strict;
use Spreadsheet::Read;
use Data::Dumper;
#http://search.cpan.org/~hmbrand/Spreadsheet-Read-0.58/Read.pm

# my $workbook  = ReadData ('Metadata-guide-record.v1.4.xls',parser => "xls");
# my $workbook  = ReadData ('Metadata-guide-record.v1.6_FF.ods',parser => "ods");
my $workbook  = ReadData ('Metadata-guide-record.v1.6_FF.xls');
my %infosOBLIGATOIRES;
my $flagInfosOBLIGATOIRESvide="none";

# print Dumper($workbook);
# print $workbook->[0]{sheet} . "\n";
InitialiseDonneesObligatoires();

my %onglets;
my  %colTOxpath;

	#on boucle sur les onglets pour remplir le hash %onglets(clef=indice de l'onglet, valeur=nom de l'onglet)
foreach (keys($workbook->[0]{sheet}))
	{
	$onglets{$workbook->[0]{sheet}{$_}}=$_;
	}

my $ongletEnCours;
my $contenuCelluleEnCours;
my $ColLgn;
my ($Col,$Lgn);

	#on boucle sur les onglets (indices croissants) pour remplir le hash %infosOBLIGATOIRES (clef=nom en dur,valeur=nom de l'onglet dans le fichier excel)
foreach $ongletEnCours(sort(keys(%onglets)))
	{
		#nom des onglets dans le fichier excel
	print "\t".$onglets{$ongletEnCours}."\n";
	$infosOBLIGATOIRES{tab_MDFields}=$ongletEnCours if ($onglets{$ongletEnCours} eq "MD Fields");
	$infosOBLIGATOIRES{tab_Help}=$ongletEnCours if ($onglets{$ongletEnCours} eq "Help");
	$infosOBLIGATOIRES{tab_MDgeneric}=$ongletEnCours if ($onglets{$ongletEnCours} eq "MD generic");
	$infosOBLIGATOIRES{tab_Thesaurus}=$ongletEnCours if ($onglets{$ongletEnCours} eq "MD Thesaurus");
	
		#identifiant des colones XPATH dans le fichier excel
	foreach $ColLgn (keys($workbook->[$ongletEnCours]))
		{
		# print $ColLgn."\n";
		if ($ColLgn=~/^([A-Z]+)([0-9]+)$/)
			{
			 ($Lgn,$Col)=($2,$1);
			$contenuCelluleEnCours=$workbook->[$ongletEnCours]{$ColLgn};
			# print $contenuCelluleEnCours."\n";
			
			if ($contenuCelluleEnCours=~/XPATH/i)
				{
				$infosOBLIGATOIRES{colone_xpath_MDgeneric}=$Col if ($ongletEnCours eq $infosOBLIGATOIRES{tab_MDgeneric});
				$infosOBLIGATOIRES{colone_xpath_Help}=$Col if ($ongletEnCours eq $infosOBLIGATOIRES{tab_Help});
				}
			}
		}
	}

 # print Dumper(%infosOBLIGATOIRES);
my $onContinue=1;
my $MessageVerifieInfosObligatoires="";
if (! VerifieInfosObligatoires($MessageVerifieInfosObligatoires))
	{
	print $MessageVerifieInfosObligatoires."\n";
	$onContinue=0;
	}
else
	{
	print "On continue\n";
	}

if ($onContinue)
	{
	coloneVersXpath();
	}

sub InitialiseDonneesObligatoires
	#procedure qui initialise à vide le hash des infos obligatoires: les noms des 4 onglets, les colones contenant les xpath
	{
	%infosOBLIGATOIRES=	(
		"tab_MDFields"				=> 	1,
		"tab_Help"					=> 	2,
		"tab_MDgeneric"				=> 	3,
		"tab_Thesaurus"				=> 	4,
		"colone_xpath_MDgeneric"		=> 	"D",
		"colone_xpath_Help"			=> 	"H",
		"colone_ids_sections"			=> 	"B",
		"1ereColoneRenseignee_MDFields"	=> 	"B",
		"1ereLigneRenseignee_Help"		=> 	4,
		"ligne_lien_xpath"			=> 	4
						);
	}

sub VerifieInfosObligatoires
	{
	my ($MonMessage)=@_;
	my $cle;
	my $InfosObligatoiresCorrect=1;
	
	foreach $cle(keys(%infosOBLIGATOIRES))
		{
		if ($infosOBLIGATOIRES{$cle} eq $flagInfosOBLIGATOIRESvide)
			{
			$MonMessage="$cle non trouve ou non rempli.\n";
			$InfosObligatoiresCorrect=0;
			}
		}
	return $InfosObligatoiresCorrect;
	}
	
sub coloneVersXpath
	{
		#on commence par remplir le hash de la ligne 4 de l'onglet 'MD Fields' %ligne4 (clef: contenu de la colone que l'on retrouve dans l'onglet help, valeur=Colone)
	my $cestNonVide=1;
	my $idCellule;
	my $iCol=$infosOBLIGATOIRES{"1ereColoneRenseignee_MDFields"};
	my %ligne4;
	
	while ($cestNonVide)
		{
		$idCellule=$iCol.$infosOBLIGATOIRES{ligne_lien_xpath};
		$contenuCelluleEnCours=$workbook->[$infosOBLIGATOIRES{tab_MDFields}]{$idCellule};
		if ($contenuCelluleEnCours)
			{
			$ligne4{$contenuCelluleEnCours}=$iCol;
			# print $iCol."\t".$contenuCelluleEnCours."\n";
			$iCol++;
			}
		else
			{
			$cestNonVide=0;
			}
		}
		#on va remplir le hash %colTOxpath (clef: numéro de colone dans l'onglet 'MD Fields' , valeur : valeur du xpath lu dans l'onglet 'Help'
	my $iLigne=$infosOBLIGATOIRES{"1ereLigneRenseignee_Help"};
	my $ya1xpath=1;
	my $idCelluleSection;
	my $section;
	while ($ya1xpath)
		{
		$idCellule=$infosOBLIGATOIRES{colone_xpath_Help}.$iLigne;
		$idCelluleSection=$infosOBLIGATOIRES{colone_ids_sections}.$iLigne;
		$contenuCelluleEnCours=$workbook->[$infosOBLIGATOIRES{tab_Help}]{$idCellule};
		$section=$workbook->[$infosOBLIGATOIRES{tab_Help}]{$idCelluleSection};
		if ($contenuCelluleEnCours)
			{
			if ($contenuCelluleEnCours=~/^\/\/*/)
				{
				$colTOxpath{$ligne4{$section}}=$contenuCelluleEnCours;
				print $ligne4{$section}."\t".$contenuCelluleEnCours."\n";
				$iLigne++;	
				}
			}
		else
			{
			$ya1xpath=0;
			}
		}
	}
