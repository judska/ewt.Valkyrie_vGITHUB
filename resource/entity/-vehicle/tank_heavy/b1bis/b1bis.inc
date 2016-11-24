
	(include "/properties/tank.ext")
	(include "/properties/shadow/tank.inc"  scale(1.2))
	(include "/properties/selection/vehicle.inc"  scale(1.2))

	{collider "tank.heavy"}
	{props "heavy" "turret_rotator" "vision_vehicle_high" "muzzle_small"}

	{extension "crew.ext"}	

	{volume "body"
		{tags "body"}
		{component "body"}
		{thickness 60}
	}
	{Volume "turret"
		{tags "turret"}
		{component "turret"}
	}
	{Volume "gun"
		{tags "gun"}
		{component "gun"}
	}
	{Volume "gun2"
		{tags "gun"}
		{component "gun2"}
	}

	("armor_heavy" turret(56 46 46 30) body(60 60 55 25))
	("armor_engine"  t(10 10 10 20))
	{sensor
		{visor "driver"
			{vision "tank_driver"}
			{bone "visor1"}
		}
		{visor "main"
			{vision "tank_main-m"}
			{bone "visor2"}
		}
		{visor "main-l"
			{vision "tank_main-l"}
			{bone "visor3"}
		}
	}
	{mass 32000}
	
	{Chassis
		("tank_trace" fx("tracks_mid") step(0.52) len(0.6))
		{Locomotion ("locomotion")
			{MaxSpeed	20}
			{maxSpeedAtMaxTurn 8.5}
			{StartTime	16}
			{TurnTime	10}
			{TurnStart  1.9} 
                        {BrakeTime         0.5}
                        {StopTime        0.7} 
    			{TurnRadius 7} 
    			{TurnStop   1}
			{Gears 0.1 0.2 0.5 0.75 1.0}
    			{RearGears 0.2} 
		}
                {speed
			{Slow	11	}
			{Normal	15}
			{Fast	20	}
		}
		{FuelBag
			{volume 600 }	; liters
			{rate	124 } 	; liters per kilometer
			{remain 600 }
			{fuel "fuel"}
		}
	}

	{Extension "b1bis.mdl"}

	{bone "turret"
		{speed 0.009}
	}

