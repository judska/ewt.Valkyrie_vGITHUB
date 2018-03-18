	(include "/properties/tank.ext")
	(include "/properties/selection/vehicle.inc" scale(1.2))
	{collider "tank.light"}
	{props "medium" "turret_rotator" "vision_vehicle_medium" "muzzle_small"}
	{PatherId "tank"}
	
	{Volume "Body1"
		{tags "body"}
		{component "body"}
	}
	{Volume "Body2"
		{tags "body"}
		{component "body"}
	}
	
	("armor_medium" turret(42 25 25 6) body(30 25 25 67))

	("crew_4_human")
	
	{boarder
		{anm "left"
			{forward	{begin	"board_tank_left"}}
			{reverse 	{end	"emit_tank_1_hold"}}
		}
		{anm "right"
			{forward	{begin	"board_tank_right"} {end "pose_mgun_stan"}}
			{reverse 	{end	"emit_tank_1_hold"}}
		}
		{anm "armor"
			{forward	{begin	"board_LowArmor"} 	{end "seat_armor"}}
			{reverse 	{end	"emit_tank_1"}}
		}
	}

	("gunner_visible")

	{extender "inventory"
		{box
			{item "ammo mgun" 2664}
			{item "bullet47 ap" 56}
			{item "bullet47 fg" 30}
			{item "breda_38" filled {user "mgun"}}
			{item "mab_38a" filled}
			{item "ammo pistol" 120}
			{item "oto_mod35 grenade" 5}
		}
	}
	{Weaponry
		("restore_ik_time")
		{place "gun"
			{RestoreIKAfterAim}
			{weapon "47mm_m47_40" filling "bullet47 ap" 1}
			{gunner "gunner"}
			{charger "commander"}
			("abm_low")
		}
		{place "mgun"
			{type "mgun_aa"}
			{foresight "foresight3"}
			{gunner "gunner2"}
			{charger "gunner2"}
			{LinkBone "mgun_rot"}
			{LinkAnimation "off"}
			("abm_mgun")
		}
		{place "mgun2"
			{type "mgun_coaxial"}
			{weapon "breda_38" filled}
			{foresight "foresight5"}
			{gunner "gunner"}
			{charger "commander"}
			("abm_mgun")
			{basic}
		}
		{place "mgun1"
			{type "mgun_hull"}
			{weapon "breda_38_x2" filled}
			{foresight "foresight4"}
			{Barrels 2}
			{gunner "gunner"}
			{charger "gunner"}
			("abm_mgun")
		}
	}	
	{mass 14300}
	{Chassis
		("tank_trace" fx("tracks_mid") step(0.52) len(0.1))
		{Locomotion ("locomotion")
			{MaxSpeed   31}
				(mod "mp"
					{maxspeed   27}
				)
			{maxSpeedAtMaxTurn 20}
			{StartTime 	 13}
			{TurnTime  	 6}
    		{TurnStart 1.4}
     		{BrakeTime   1.5} 
    		{StopTime    2}
     		{TurnRadius  8} 
    		{TurnStop   0.7}
     		{Gears  0.2 0.4 1.0}
     		{RearGears 0.1 0.45}
		}
		(mod "mp"
			{FuelBag
				{volume 320 }	
				{rate	145 }   
				{remain 320 }
				{fuel "fuel"}
			}
		    {speed
        		{Normal  12}
    		}
		)
	}
	{Extension "m1441.mdl"}
	{bone "turret"
		 {speed 0.010}
	}
	{bone "gun_rot"
		 {speed 0.002}
	}


