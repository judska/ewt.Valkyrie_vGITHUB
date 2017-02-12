
	(include "/properties/tank.ext")
	(include "/properties/shadow/tank.inc"  scale(1.2))
	(include "/properties/selection/vehicle.inc"  scale(1.2))
	{collider "tank.heavy"}
	{props "heavy" "turret_rotator" "vision_vehicle_high" "muzzle_medium"}

	{extender "enumerator"}


	{Volume "body1"
		{tags "body"}
		{component "body"}
	}
	{Volume "body2"
		{tags "body"}
		{component "body"}
	}
	{Volume "body3"
		{tags "body"}
		{component "body"}
	}


	("armor_engine"  t(10 95 51 20))

	("shield" args 8)
	{extender "shield"
		{on}
	}

	{component "shieldl1" {tags "shield"}}
	{component "shieldl2" {tags "shield"}}
	{component "shieldr1" {tags "shield"}}
	{component "shieldr2" {tags "shield"}}

	
	
	("crew_5_human")

		{boarder
		{anm "left"
			{forward	{begin	"board_tank_left"}}
			{reverse 	{end	"emit_tank_1_hold"}}
		}
		{anm "right"
			{forward	{begin	"board_tank_right"}}
			{reverse 	{end	"emit_tank_1_hold"}}
		}
		{anm "armor"
			{forward	{begin	"board_lowArmor"} 	{end "pose_seat_armor_2"}}
			{reverse 	{end	"emit_tank_2"}}
		}
		{anm "armor1"
			{forward	{begin	"board_lowArmor"} {end "pose_seat_armor_2"}}
			{reverse 	{end	"emit_body"}}
		}
		{anm "armor2"
			{forward	{begin	"board_lowArmor"} {end "pose_seat_armor_1"}}
			{reverse 	{end	"emit_body"}}
		}
	}
	
	("seat_human" place(seat1) bone(seat01) door(emit3)  animation(armor1))
	("seat_human" place(seat2) bone(seat02) door(emit4)  animation(armor))
	("seat_human" place(seat3) bone(seat03) door(emit5)  animation(armor))
	("seat_human" place(seat4) bone(seat04) door(emit6)  animation(armor1))





