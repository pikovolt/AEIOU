
// リマップ処理
function apply_timeRemap(sKeylist)
{
	// JSON文字列から、オブジェクト生成 (※エラーチェックとか、なし..)
	var keylist = eval( '(' + sKeylist + ')' );
	if(keylist.property != 'Time Remap')
	{
		alert("指定プロパティは、'Time Remap'のみのサポートになっています.");
		return;
	}

	// レイヤー選択のチェック (ScaleSelected Layer.jsxのパクリ.)
	var activeItem = app.project.activeItem;
	if ((activeItem == null) || !(activeItem instanceof CompItem))
	{
		alert("コンポジションを開いてから、スクリプトを実行してください.");		// Please select or open a composition first.
		return;
	}
	var selectedLayers = activeItem.selectedLayers;
	if (activeItem.selectedLayers.length == 0)
	{
		alert("レイヤーを選択してください.(複数選択:可)");	// Please select at least one layer in the active comp first.
		return;
	}

	// アンドゥ対象の処理：開始位置
	app.beginUndoGroup("apply_timeRemap");

	// 選択されたレイヤー(複数選択可)に、リマップ情報を適用
	for(var num in activeItem.selectedLayers)
	{
		// レイヤーを取得
		var activeLayer = activeItem.selectedLayers[num];

		// リマップ情報を初期化(リマップ済みの場合もあるので、一回リマップ解除、その後リマップ適用の上で、最終フレームのキーを削除)
		activeLayer.timeRemapEnabled = false;
		activeLayer.timeRemapEnabled = true;
		activeLayer.property("Time Remap").removeKey(2);

		// リマップ情報を設定
		var firstKeyUsed = false;
		for(var k in keylist.keys)
		{
			// 関数の引数として受け取った情報を、選択レイヤーに設定
			var key = keylist.keys[k];
			var t = key.t    / keylist.scale;
			var v = key.v[0] / keylist.scale;
			try{
				activeLayer.property("Time Remap").setValueAtTime(t, v);
			}
			catch(e){
				// 設定できない時刻への反映はスキップ
			}

			// 先頭フレームにキーを打っているかをチェック（打った場合はフラグを立てる.）
			if(key.t == 0.0)
			{
				firstKeyUsed = true;
			}
		}
		// 先頭フレームにキーを打たなかった場合、先頭のキーを削除
		//(※都合上、先頭より前にキーがあると不味い…のだが、打てるんだっけか？)
		if(!firstKeyUsed)
		{
			activeLayer.property("Time Remap").removeKey(1);
		}

		//全てのキーを"HOLD"に設定
		var count = activeLayer.property("Time Remap").numKeys;
		for(var i = 1; i <= count; i++)
		{
			activeLayer.property("Time Remap").setInterpolationTypeAtKey(i,KeyframeInterpolationType.HOLD,KeyframeInterpolationType.HOLD);
		}
	}

	// アンドゥ対象の処理：終了位置
	app.endUndoGroup();
}

// ClipBoardから文字列取得
var sKeylist = system.callSystem("D:\\getClipBoard.exe");

// リマップ適用
apply_timeRemap(sKeylist);
