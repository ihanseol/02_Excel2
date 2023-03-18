



' W-1~~ sigma  _{w-1} = {2 pi  TIMES  1.732 TIMES  3.99} over {130} -1.15 TIMES  log {2.25 TIMES  1.732 TIMES  (1/1440)} over {0.0005 TIMES  (0.100 TIMES  0.100)} =`-2.8095




' in here every variable are string
"W-" & well & "~~ sigma  _{w-" & well & "} = {2 pi  TIMES  " & T & " TIMES  " & delta_s & " } over {" & Q & "} -1.15 TIMES  log {2.25 TIMES  " & T & " TIMES  (1/1440)} over {0.0005 TIMES  (" & radius & " TIMES  " & radius & ")} =`" & skin_factor


' W-1~~r _{e-1} `=~r _{w} e ^{- sigma  _{w-1}} =0.100 TIMES ℯ ^{-(-2.8095)} =1.6600m
"W-" & well & "~~r _{e-" well & "} `=~r _{w} e ^{- sigma  _{w-" & well & "}} =" & radius & " TIMES ℯ ^{-(" & skin_factor & ")} =" & er & "m"


' W-1~~Q _{2} `＝149` TIMES  `(` {44.94} over {21.53} `) ^{2/3} `＝134.0㎥/일
"W-" & {well} & "~~Q _{ & 2} `＝" & Q & "` TIMES  `(` {" & Q1 & "} over {" & Q2 & "} `) ^{2/3} `＝" & res & "㎥/일"






' W-1호공~~R _{W-1} ``=`` sqrt {6 TIMES  50.5 TIMES  0.00338 TIMES  0.5833/0.0003229} ``=~43.0m

"W-" & well & "호공~~R _{W-" & well & "} ``=`` sqrt {6 TIMES  " & delta_h & " TIMES  " & K & " TIMES  " & time & "/" & S & "} ``=~" & schultze & "m"




' W-1호공~~R _{W-1} ``=``3 sqrt {50.5 TIMES  0.00338 TIMES  0.5833/0.0003229} `=`52.7`m

"W-" & well & "호공~~R _{W-" & well & "} ``=``3 sqrt {" & delta_h & " TIMES  " & K & " TIMES  " & time & "/" & S "} `=`" & weber & "`m"


' W-1호공~~r _{0(W-1)} `=~ sqrt {{2.25 TIMES  1.5550 TIMES  0.5833} over {0.0003229}} `=~81.7m



"W-" & well & "호공~~r _{0(W-" & well & ")} `=~ sqrt {{2.25 TIMES  " & T & " TIMES  " & time & "} over {" & S & "}} `=~" & jcob & "m"