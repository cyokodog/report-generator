#Report Generator

##�T�v

Report Generator �̓G�N�Z���`���̒��[�o�̓v���O�����������������邽�߃c�[���ł��B���̃c�[���y�сA�������������v���O������ VBScript �x�[�X�̂��� Windows ���ł̂ݎg�p���\�ł��BIE ����ɂȂ�܂��� ActiveX �̎g�p���L���Ȋ��ł���� Web �A�v���P�[�V�����ɑg�ݍ��ނ��Ƃ��ł��܂��B

�ȉ��̂悤�Ȏ菇�ō쐬���܂��B

1. �e���v���[�g�̒��[�G�N�Z���t�@�C���̍쐬
2. �e���v���[�g�t�@�C���ւ̃f�[�^���ʎq�̋L�q
3. Report Generator �ɂ�钠�[�o�̓v���O�����̐���
4. �A�v���P�[�V�����ւ̑g�ݍ���

##�`���[�g���A��

�ȉ���� Report Generator ���_�E�����[�h���𓀂��܂��B

[cyokodog / report-generator | GitHub](https://github.com/cyokodog/report-generator)

�𓀂���Ɖ��L�\���̃t�H���_�A�t�@�C�����ł�������܂��B([ ]�̓t�H���_)

	-[ReportGenerator]
		-ReportGenerator-*.*.*.wsf
		-[lib]
			-generateReport.vbs
			-init.vbs
			-jquery-1.4.2.min
		-[sample_app]
			-[ex01]
			-[ex02]
			�E�E�E
		-[sample_template]
			-mitsumori.xls
			-mitsumori_map.xls
			-mitsumori_sample.xls

###�e���v���[�g�̒��[�G�N�Z���t�@�C���̍쐬

�܂��A�e���v���[�g�ƂȂ钠�[�G�N�Z���t�@�C�����쐬���܂��Bsample_template �t�H���_�� mitsumori.xls �̂悤�ɍ��܂��B

**mitsumori.xls**

![mitsumori.xls](http://cdn-ak.f.st-hatena.com/images/fotolife/c/cyokodog/20100921/20100921023813.png)

�K�v�ɉ����e�Z���ɑ΂������⏑������ݒ肵�܂��B

###�e���v���[�g�t�@�C���ւ̃f�[�^���ʎq�̋L�q

�σf�[�^�𖄂ߍ��މӏ��Ɏ��ʎq���L�q���A�t�@�C�����̖����� "_map" �ƕt���ăt�@�C����ۑ����܂��Bsample_template �t�H���_�� mitsumori_map.xls �̂悤�ɋL�q���܂��B

**mitsumori_map.xls**

![mitsumori_map.xls](http://cdn-ak.f.st-hatena.com/images/fotolife/c/cyokodog/20100921/20100921023840.png)

���ʎq�� {�f�[�^��} �Ƃ����`���ŋL�q���܂��B�Ⴆ�Δ[�i�����Ȃ� {delivery_date} �̂悤�ɂ��ċL�q���܂��B

###Report Generator �ɂ�钠�[�o�̓v���O�����̐���

mitsumori_map.xls �� ReportGenerator �t�H���_�� ReportGenerator-*.*.*.wsf �ɑ΂��h���b�v���܂��B

![drop to ...](http://cdn-ak.f.st-hatena.com/images/fotolife/c/cyokodog/20100921/20100921024637.png)

�v���O���������������n�܂�̂ŁA�����I���̃��b�Z�[�W���\�������܂ŏ����҂��܂��B

�������������邷��Ɖ��L�\���̃t�H���_�A�t�@�C�����ł�������܂��B([ ]�̓t�H���_)

	-[result_mitsumori]
		-mitsumori.vbs
		-mitsumori.xls
		-mitsumori.xml
		-[bat_sample]
			-mitsumori.bat
			-mitsumori.wsf
			-mitsumori_dat.xml
		-[html_sample]
			-jquery-1.4.2.min.js
			-mitsumori.html
			-mitsumori_dat.xml

�E  
�E  
�E  
�E  
�E  
�E  
�E  
��ꂽ�E�E�E

������[���o�[�W�����̃h�L�������g](http://d.hatena.ne.jp/cyokodog/20100927/reportgenerator01)�����Q�Ƃ��������B��{�I�ɂ��Ƃ͕ς���Ă܂���̂ŁE�E�E��������������ɂ��ǋL���Ă��܂��B

���A���� xlsx �ɑΉ�������ꍇ�́Alib �t�H���_�� init.vbs ���̋L�q���ȉ��̂悤�ɕύX���Ă��������B

	Const XLS_EXT = ".xlsx"





