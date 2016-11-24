	(include "/properties/armoredcar.ext")
	(include "/properties/shadow/tank.inc"  scale(1.0))
	(include "/properties/selection/vehicle.inc"  scale(1.0))
	{props "vehicle" "armored_car" "tank" "turret" "wheeled" "vision_vehicle_low"}
	{targetClass "armoredcar" "ap bullet"}
	{targetSelector "tank"}
	(include "/properties/tank_crew.ext")
	{PatherId "vehicle.small"}

	{extender "enumerator"}
	
	{volume "body1"
		{tags "body"}
		{component "body"}
		{thickness 8
		    {front 14}
		}
	}
	{volume "body2"
		{tags "body"}
		{component "body"}
		{thickness 8
		    {front 14}
		}
	}
	{volume "body3"
		{tags "body"}
		{component "body"}
		{thickness 8
		    {front 14}
		}
	}
	{volume "body4"
		{tags "body"}
		{component "body"}
		{thickness 8
		    {front 14}
		}
	}
	{volume "body5"
		{tags "body"}
		{component "body"}
		{thickness 8
		    {front 14}
		}
	}
	{volume "body6"
		{tags "body"}
		{component "body"}
		{thickness 8
		    {front 14}
		}
	}
	{volume "body7"
		{tags "body"}
		{component "body"}
		{thickness 8
		    {front 14}
		}
	}
	{volume "turret"
		{tags "turret"}
		{component "turret"}
		{thickness 7}
	}
	{volume "turret2"
		{tags "turret"}
		{component "turret"}
		{thickness 7}
	}
	{volume "turret3"
		{tags "turret"}
		{component "turret"}
		{thickness 7}
	}
	{volume "turret4"
		{tags "turret"}
		{component "turret"}
		{thickness 7}
	}
	{volume "turret5"
		{tags "turret"}
		{component "turret"}
		{thickness 7}
	}
	{volume "turret6"
		{tags "turret"}
		{component "turret"}
		{thickness 7}
	}
	{volume "roofshield"
		{tags "roofshield"}
		{component "roofshield"}
		{thickness 7}
	}
	{volume "engine"
		{tags "engine"}
		{component "engine"}
		{Thickness 7
			{front 14}
		}
	}
	
	

	("armor_medium" args 1.2)
	
	{component "turret"
		{RepairTime 20}
	}

	{boarder
		{anm "driver"
			{forward	{begin	"board_daimler_gunner"} {end "seat_driver_car"} {base "doorright"}}
			{reverse 	{end	"board_daimler_gunner" -1} {base "doorright" -1}}
		}
		{anm "gunner"
			;{forward	{begin	"board_daimler_gunner"} {end "seat_m7_mgunner"} {base "doorleft"}}
			{forward	{begin	"board_daimler_gunner"} {end "seat_sd223_gunner"} {base "doorleft"}}
			{reverse 	{end	"board_daimler_gunner" -1} {base "doorleft" -1}}
		}
		{anm "charger"
			{forward	{begin	"board_daimler_gunner"} {end "seat_driver_car"} {base "doorright"}}
			{reverse 	{end	"board_daimler_gunner" -1} {base "doorright" -1}}
		}
		{anm "commander"
			{forward	{begin	"board_daimler_gunner"} {end "seat_driver_car"} {base "doorleft"}}
			{reverse 	{end	"board_daimler_gunner" -1} {base "doorleft" -1}}
		}
		
	}

	(define "crew_human"
	{placer
		{place "%place" {group "%group"} {visor "%visor"}}
		{place "gunner"	{LinkBone "gunner"}	{Group "crew"}	{turnoff {sensor} {shadow}} {visor "main"}}	
	}
	{boarder
		{door "%door"}
		{link "%door" "%place" {anm "%animation"} {forward putoff} {reverse puton}}
	}
)
	
	("crew_human" place(driver)		door(emit2)	group(crew)	animation(driver)	visor(driver))
	("crew_human" place(charger)		door(emit2)	group(crew)	animation(charger)	visor(driver))
	("crew_human" place(commander)	door(emit1)	group(crew)	animation(commander)	visor(driver))
	("crew_human" place(gunner) 		door(emit1)	group(crew)	animation(gunner)	visor(main))
	{placer {place "gunner" {openForCrush 0}}}
			
			

	{extender "inventory"
		{box
			{item "ammo mgun" 800}
			{item "ammo ptr" 60}
		}
	}
	
	{Weaponry
		{place "mgun"
			{weapon "boyes" filled}
			{foresight "foresight3"}
			{LinkBone "boyes"}
			{LinkAnimation "off"}
			{charger "gunner"}
			("abm_low")
		}
		{place "mgun1"
			{weapon "bren" filled}
			{foresight "foresight4"}
			{type "mgun_aa"}
			{LinkBone "mgun1"}
			{charger "gunner"}
			{LinkAnimation "off"}
			("abm_mgun")
			{basic}
		}
	}

	("wheel" args "wheelright1")
	("wheel" args "wheelleft1")
	("wheel" args "wheelright2")
	("wheel" args "wheelleft2")

	{mass 4800}
	{Chassis
		{Locomotion ("locomotion")
			{MaxSpeed	45}
			(mod "mp" 
				{MaxSpeed  41}
			)
	  		{maxSpeedAtMaxTurn 15}		
			{StartTime	11}
			{BrakeTime	1.5}	
			{StopTime	3}
			{TurnRadius	5}
		}
		{FuelBag
			{volume 90}
			{rate	2.3} 
			{remain 90 }
			{fuel "fuel"}
		}
		("car_trace" fx("tracewheel02") step(0.09) wheell("wlf") wheelr("wrf") bonel("FXtraceL1") boner("FXtraceR1"))
		("car_trace" fx("tracewheel02") step(0.09) wheell("wlr") wheelr("wrr") bonel("FXtraceL2") boner("FXtraceR2"))
	}
	{Extension "morriscs9.mdl"}
	{bone "turret"
		{speed 0.010}
	}

