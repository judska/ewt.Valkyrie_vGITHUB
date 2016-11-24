{"barbed_wire"
	{on "cut_pieces"
		{start_sound "crush/small/wire"}
	}
	{on break_armor
		{if name "blast"
			{if min_energy 25
				{call "cut_pieces"}
			}
		}
	}
	{on bullet_hit
		{if stuff "shell"
			{call "cut_pieces"}
			{if volume "piece"
				{if stuff "fg"
					{bullet_detonate}
					{call "cut_pieces"}
				else stuff "grenade"
					{bullet_detonate}
					{call "cut_pieces"}
				}
			}
		}
	}
}
