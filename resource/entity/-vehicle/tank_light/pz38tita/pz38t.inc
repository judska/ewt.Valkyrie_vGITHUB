	(include "/properties/tank.ext")
	(include "/properties/selection/vehicle.inc" scale(1.0))
	{props "light" "turret" "vision_vehicle_low" "muzzle_small"}
	
	{collider "tank.light"}
	{targetClass "light_tank"}
	{patherId "tank.light"}
	
	{extender "enumerator"
		{digit_folder "ger2"}
	}
	
	{volume "body1"
		{tags "body"}
		{component "body"}
	}
	{volume "body2"
		{tags "body"}
		{component "body"}
	}
	{volume "body3"
		{tags "body"}
		{component "body"}
	}
	{volume "body4"
		{tags "body"}
		{component "body"}
	}

	("crew_4_human_extended")
	("seat_4_human")

	{boarder
		{anm "driver"
			{forward	{begin	"board_pz4_driver"} {end "seat_gunner_stand"} {base "open_driver" -0.4}}
			{reverse 	{end	"emit_tank_2_hold"} {base "open_driver" 0.9}}
		}
	}

	{extender "inventory"
		{box
			{item  "ammo mgun" 2400}
			{item "bullet37 ap" 55}
			{item "bullet37 apcr" 5}
			{item "bullet37 fg" 30}	
		}
	}
	{Weaponry
		("restore_ik_time")
		{place "gun"
			{weapon "37mm_pak36" filling "bullet37 ap"}
			{RestoreIKAfterAim}
			{gunner "commander"}
			{charger "charger"}
			("abm_low")
		}
		{place "mgun"
			("abm_mgun")
			{type "mgun_coaxial"}
			{weapon "mg37" filled}
			{foresight "foresight3"}
			{gunner "commander"}
			{charger "charger"}
			{basic}
			("abm_mgun")
		}
		{place "mgun1"
			("abm_mgun")
			{type "mgun_hull"}
			{weapon "mg37" filled}
			{foresight "foresight4"}
			{gunner "driver"}
			{charger "driver"}
			("abm_mgun")
		}
	}

	{mass 9400}
       	 {sensor
		{visor "main"
			{vision "tank_main-l"}
		}
		{visor "main1"
			{vision "tank_main-m"}
			{bone "foresight1"}
		}
		{visor "view-around"
			{vision "vehicle_top_turret"}
			{bone "turret"}
		}
	}

	{mass 9850}
	{Chassis
		("tank_trace" fx("tracks_mid") step(0.56) len(0.3))
		{Locomotion ("locomotion")
			{maxspeed   24}
			{maxSpeedAtMaxTurn 5}
			{StartTime  10}
			{TurnStart  2}			
			{TurnTime   5}
			{TurnRadius 10} 
    			{TurnStop   0.9}
			{BrakeTime  1}
			{StopTime   2}
			{Gears      0.1 0.2 0.3 0.5 1}
			{RearGears  0.3}		
		}
		{FuelBag
			{volume 220 }	
			{rate	1.4 }   
			{remain 220 }
			{fuel "fuel"}
		}
		{speed
        		{Normal  18}
    		}
	}

	{Extension "pz38t.mdl"}

	{bone "turret"
		{speed 0.009}
	}
	{bone "gun_rot"
		{speed 0.02}
	}	


