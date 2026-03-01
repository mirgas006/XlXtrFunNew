Toto je náhrada za pôvodné rozšírenie Excelu XlXtrFun.
Nahradené sú všetky 1D a 2D funkcie.
	LookupClosestValue, IndexOfClosestValue, LookupClosestValue2D, Spline, Interp a Interpolate, PFit fungujú identicky.
	Pri funkciách derivácií dydx a ddydx sú odchýlky, pri dydx zanedbateľné pri ddydx v miestach s prudkou zmenou významnejšie.
		Súbor XlXtrFun-alternativne-derivacie.py obsahuje alternatívnu logiku derivovania, od pôvodnej sa taktiež líši (porovnateľne ako prednastavená),
			pre jej impelmentáciu treba nahradiť súbor "XlXtrFun.py" v adresári "python" premenovaným alternatívnym súborom.
	Funkcia PFitData na novších verziách Excelu s pôvodným rozšírením už nefungovala. V novom rozšírení je plne funkčná.
	Pri funkcii Intercept sú výsledky identické avšak nová verzia niekedy potrebuje vyšší počet iterácií
		=> odporučenie, pri chybe skúsiť zvýšiť počet Max_Iterations
	Funkcia XatY dáva identické výsledky, avšak je citlivejšia na hodnotu GuessX. Oproti pôvodnej verzii pri nájdení extrému overí, či ide o hladané minimum/maximum alebo inflex.
		Pôvodná verzia vždy vrátila hodnotu extrému bez ohľadu na druhú deriváciu, pri hľadaní minima mohla vrátiť maximum a naopak.

V súbore "porovnanie.xlsx" je možné nahliadnuť na rozdieli v pôvodnej a tejto verzii.

Adresár "source" obsahuje zdrojový kód rozšírenia.
Adresár "dist" obsahuje skompilované rozšírenie pre použivateľa.
V prípade, že Excel po pridaní rozšírenia do Excelu štrajkuje, treba rozšírenie .xll "odblokovať", tak ako to je spravené na obrázku "permission.png".