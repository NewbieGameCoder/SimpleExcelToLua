return {
	ID = 11037920, --列头1(在有“|”这个标志的时候，如果读数据默认只读一行有效数据)
	ID2 = {11037920, 11037999, 11038000, 11038001, 11038002},--列头2
	ID3 = --列头3
	{
		{
			11037920, --注释
			11037999, --注释
			11038000, --注释
			11038001, --注释
		},
	},
	ID4 = --列头4
	{
		{11037920},--注释

		{11037999},--注释

		{11038000},--注释

		{11038001},--注释

		{11038002},--注释

	},
	ID5 = --列头5(代表只有3行有效数据)
	--Key代表注释2
	{
		[1] = 
			"1", --注释3
			--Key代表注释4
			{
				[1] = 
					1, --注释5

			},
		[2] = 
			"2", --注释3
			--Key代表注释4
			{
				[2] = 
					2, --注释5

			},
		[3] = 
			--Key代表注释4
			{
				[3] = 
					3, --注释5

			},
	},
	ID6 = --列头6
	{
		{
			TestNest1 = "1", --注释6
			NestEle2 = 1, --注释7
			NestEle3 = 
				{
					1, --注释8
					1, --注释9
				},
			TestNest2 = 
				--Key代表注释10
				{
					["1"] = 
						{
							NestEle2 = 1, --注释11
							NestEle3 = 1, --注释12
						},
				},
		},
		{
			TestNest1 = "2", --注释6
			NestEle2 = 2, --注释7
			NestEle3 = 
				{
					2, --注释8
					2, --注释9
				},
			TestNest2 = 
				--Key代表注释10
				{
					["2"] = 
						{
							NestEle2 = 2, --注释11
							NestEle3 = 2, --注释12
						},
				},
		},
		{
			TestNest1 = "3", --注释6
			NestEle2 = 3, --注释7
			NestEle3 = 
				{
					3, --注释8
					3, --注释9
				},
			TestNest2 = 
				--Key代表注释10
				{
					["3"] = 
						{
							NestEle2 = 3, --注释11
							NestEle3 = 3, --注释12
						},
				},
		},
		{
			TestNest1 = "4", --注释6
			NestEle2 = 4, --注释7
			NestEle3 = 
				{
					4, --注释8
					4, --注释9
				},
			TestNest2 = 
				--Key代表注释10
				{
					["4"] = 
						{
							NestEle2 = 4, --注释11
							NestEle3 = 4, --注释12
						},
				},
		},
	},
}