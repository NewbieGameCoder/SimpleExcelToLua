return {
	{
		ID = 11037920, --ID,
		Description = "测试", --字符串,
		TestNest1 = 
			{
				1, --注释1,
				1, --注释2,
			},
		TestNest2 = 
			{
				"1", --注释3,
				--Key代表注释4
				{
					[1] = 
						1, --注释5,
				},
			},
		TestNest3 = 
			--Key代表注释6
			{
				["1"] = 
					{
						NestEle2 = 1, --注释7,
						NestEle3 = 
							{
								1, --注释8,
								1, --注释9,
							},
					},
			},
		TestNest4 = 
			--Key代表注释10
			{
				["1"] = 
					{
						NestEle2 = 1, --注释11,
						NestEle3 = 1, --注释12,
					},
			},
	},
	{
		ID = 11037999, --ID,
		Description = {"\"测试1\"", "\"测试2\""}, --字符串,
		TestNest1 = 
			{
				2, --注释1,
				2, --注释2,
			},
		TestNest2 = 
			{
				"2", --注释3,
				--Key代表注释4
				{
					[2] = 
						2, --注释5,
				},
			},
		TestNest3 = 
			--Key代表注释6
			{
				["2"] = 
					{
						NestEle2 = 2, --注释7,
						NestEle3 = 
							{
								2, --注释8,
								2, --注释9,
							},
					},
			},
		TestNest4 = 
			--Key代表注释10
			{
				["2"] = 
					{
						NestEle2 = 2, --注释11,
						NestEle3 = 2, --注释12,
					},
			},
	},
	{
		ID = 11038000, --ID,
		TestNest2 = 
			{
				--Key代表注释4
				{
					[3] = 
						3, --注释5,
				},
			},
		TestNest3 = 
			--Key代表注释6
			{
				["3"] = 
					{
						NestEle2 = 3, --注释7,
						NestEle3 = 
							{
								3, --注释8,
								3, --注释9,
							},
					},
			},
		TestNest4 = 
			--Key代表注释10
			{
				["3"] = 
					{
						NestEle2 = 3, --注释11,
						NestEle3 = 3, --注释12,
					},
			},
	},
	{
		ID = 11038001, --ID,
		Description = "测试\"字符串\"", --字符串,
		TestNest1 = 
			{
				4, --注释1,
				4, --注释2,
			},
		TestNest2 = 
			{
				"4", --注释3,
				--Key代表注释4
				{
					[4] = 
						4, --注释5,
				},
			},
		TestNest3 = 
			--Key代表注释6
			{
				["4"] = 
					{
						NestEle2 = 4, --注释7,
						NestEle3 = 
							{
								4, --注释8,
								4, --注释9,
							},
					},
			},
		TestNest4 = 
			--Key代表注释10
			{
				["4"] = 
					{
						NestEle2 = 4, --注释11,
						NestEle3 = 4, --注释12,
					},
			},
	},
	{
		ID = 11038002, --ID,
		Description = "测试“字符串”", --字符串,
		TestNest1 = 
			{
				5, --注释1,
				5, --注释2,
			},
		TestNest2 = 
			{
				"5", --注释3,
				--Key代表注释4
				{
					[5] = 
						5, --注释5,
				},
			},
		TestNest3 = 
			--Key代表注释6
			{
				["5"] = 
					{
						NestEle2 = 5, --注释7,
						NestEle3 = 
							{
								5, --注释8,
								5, --注释9,
							},
					},
			},
		TestNest4 = 
			--Key代表注释10
			{
				["5"] = 
					{
						NestEle2 = 5, --注释11,
						NestEle3 = 5, --注释12,
					},
			},
	},
}