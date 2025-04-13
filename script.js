// テストデータ（Excelから抽出した情報）
const testData = [
    { level: 1, item: "売り場", partOfSpeech: "noun", correctAnswer: "department", distractors: ["boy", "meeting", "town"], discrimination: 1.98156526596195, difficulty: -2.38498224191404, guessing: 0.130828033790051 },
    { level: 1, item: "願う", partOfSpeech: "verb", correctAnswer: "wish", distractors: ["return", "take", "feel"], discrimination: 4.14847237503316, difficulty: -1.85306682792874, guessing: 0.0576237773032269 },
    { level: 1, item: "許可する", partOfSpeech: "verb", correctAnswer: "let", distractors: ["try", "buy", "spend"], discrimination: 1.65148933661613, difficulty: -1.18471877186228, guessing: 0.0160932143473078 },
    { level: 1, item: "丸い", partOfSpeech: "adjective", correctAnswer: "round", distractors: ["ready", "simple", "easy"], discrimination: 2.17501052164738, difficulty: -2.40345043459572, guessing: 0.0465664027259637 },
    { level: 1, item: "全体の", partOfSpeech: "adjective", correctAnswer: "whole", distractors: ["important", "other", "large"], discrimination: 3.83352007532768, difficulty: -0.946813602465208, guessing: 0.25271458704626 },
    { level: 1, item: "普通は", partOfSpeech: "adverb", correctAnswer: "usually", distractors: ["never", "always", "finally"], discrimination: 0.855750317619688, difficulty: -3.07154361283321, guessing: 0.0595256489100603 },
    { level: 1, item: "地域", partOfSpeech: "noun", correctAnswer: "area", distractors: ["style", "art", "record"], discrimination: 3.32915563467688, difficulty: -2.47167122416048, guessing: 0.0238615725582434 },
    { level: 1, item: "銀行", partOfSpeech: "noun", correctAnswer: "bank", distractors: ["subject", "age", "century"], discrimination: 0.866970064751077, difficulty: -3.54173003220117, guessing: 0.0627346680788252 },
    { level: 1, item: "実行", partOfSpeech: "noun", correctAnswer: "performance", distractors: ["project", "mind", "president"], discrimination: 1.3986090190306, difficulty: 1.00199336043642, guessing: 0.151417863401701 },
    { level: 1, item: "１年", partOfSpeech: "noun", correctAnswer: "year", distractors: ["state", "power", "father"], discrimination: 0.917448900354817, difficulty: -3.98862771189938, guessing: 0.0687793602611114 },
    { level: 1, item: "特別の", partOfSpeech: "adjective", correctAnswer: "special", distractors: ["western", "key", "old"], discrimination: 6.78654470508513, difficulty: -2.2540690242761, guessing: 0.0194054357487844 },
    { level: 1, item: "それから", partOfSpeech: "adverb", correctAnswer: "then", distractors: ["alone", "else", "far"], discrimination: 0.85843402353509, difficulty: -2.18476071331574, guessing: 0.049001429087084 },
    { level: 1, item: "押すこと", partOfSpeech: "noun", correctAnswer: "pressure", distractors: ["field", "support", "question"], discrimination: 1.35383960528495, difficulty: -2.04312483977279, guessing: 0.101238783759803 },
    { level: 1, item: "始める", partOfSpeech: "verb", correctAnswer: "begin", distractors: ["stand", "do", "want"], discrimination: 1.53094369523002, difficulty: -2.97375117121192, guessing: 0.048961249896083 },
    { level: 1, item: "過去の", partOfSpeech: "adjective", correctAnswer: "past", distractors: ["major", "each", "red"], discrimination: 5.33520026596577, difficulty: -2.119936872872, guessing: 0.0198152121505737 },
    { level: 1, item: "部屋", partOfSpeech: "noun", correctAnswer: "room", distractors: ["daughter", "increase", "front"], discrimination: 1.29966859265209, difficulty: -2.86663852148445, guessing: 0.0629648385109624 },
    { level: 1, item: "進路", partOfSpeech: "noun", correctAnswer: "course", distractors: ["group", "music", "order"], discrimination: 1.61528309357791, difficulty: -2.28427668892411, guessing: 0.0906774561460272 },
    { level: 1, item: "玄関", partOfSpeech: "noun", correctAnswer: "hall", distractors: ["size", "place", "member"], discrimination: 1.3760533015799, difficulty: -1.69717237095458, guessing: 0.0813400441251794 },
    { level: 1, item: "小休止", partOfSpeech: "noun", correctAnswer: "break", distractors: ["heart", "hour", "bed"], discrimination: 1.39897613849308, difficulty: -1.63834703102039, guessing: 0.0650684134362941 },
    { level: 1, item: "持つ", partOfSpeech: "verb", correctAnswer: "hold", distractors: ["call", "remember", "bring"], discrimination: 1.46749296509345, difficulty: -3.2131180717206, guessing: 0.0762282660230694 },
    { level: 2, item: "…だと思う", partOfSpeech: "verb", correctAnswer: "suppose", distractors: ["lay", "succeed", "gain"], discrimination: 2.13111385571391, difficulty: -1.19438284852792, guessing: 0.157694506947415 },
    { level: 2, item: "たぶん", partOfSpeech: "adverb", correctAnswer: "perhaps", distractors: ["quite", "similarly", "effectively"], discrimination: 1.61905167386398, difficulty: -1.12574011058146, guessing: 0.0433657811915526 },
    { level: 2, item: "請求書", partOfSpeech: "noun", correctAnswer: "bill", distractors: ["society", "scene", "measure"], discrimination: 2.47470191298633, difficulty: -1.09895373691181, guessing: 0.00495584891297062 },
    { level: 2, item: "範囲", partOfSpeech: "noun", correctAnswer: "range", distractors: ["plenty", "instrument", "surface"], discrimination: 2.02976456761161, difficulty: -1.75584445947967, guessing: 0.0283425473282377 },
    { level: 2, item: "一般的な", partOfSpeech: "adjective", correctAnswer: "general", distractors: ["recent", "empty", "illegal"], discrimination: 2.03010085188381, difficulty: -1.73208741126293, guessing: 0.0493790289957487 },
    { level: 2, item: "地域", partOfSpeech: "noun", correctAnswer: "region", distractors: ["silver", "content", "customer"], discrimination: 2.77862023765874, difficulty: -1.14657244002597, guessing: 0.222231583183946 },
    { level: 2, item: "構造", partOfSpeech: "noun", correctAnswer: "structure", distractors: ["neighborhood", "lack", "application"], discrimination: 3.50019614030359, difficulty: -1.24512418908157, guessing: 0.147122778806464 },
    { level: 2, item: "制御", partOfSpeech: "noun", correctAnswer: "control", distractors: ["benefit", "stomach", "supply"], discrimination: 2.79722394549621, difficulty: -1.93500493480634, guessing: 0.018540893169078 },
    { level: 2, item: "型", partOfSpeech: "noun", correctAnswer: "type", distractors: ["discussion", "string", "gas"], discrimination: 6.44931576796345, difficulty: -2.18539993439464, guessing: 0.0211693306776987 },
    { level: 2, item: "患者", partOfSpeech: "noun", correctAnswer: "patient", distractors: ["technique", "user", "aid"], discrimination: 1.97299204452406, difficulty: -1.67298786376143, guessing: 0.0537576013748299 },
    { level: 2, item: "苦痛", partOfSpeech: "noun", correctAnswer: "pain", distractors: ["website", "appeal", "solution"], discrimination: 3.0210187035265, difficulty: -1.2110156466258, guessing: 0.250804884473934 },
    { level: 2, item: "よく考える", partOfSpeech: "verb", correctAnswer: "consider", distractors: ["overcome", "preserve", "approve"], discrimination: 2.39924989677358, difficulty: -0.923586017514527, guessing: 0.0858403680604062 },
    { level: 2, item: "利用可能な", partOfSpeech: "adjective", correctAnswer: "available", distractors: ["central", "everyday", "huge"], discrimination: 3.73463141032276, difficulty: -1.6217202884235, guessing: 0.083468228565789 },
    { level: 2, item: "最近", partOfSpeech: "adverb", correctAnswer: "recently", distractors: ["somewhere", "newly", "heavily"], discrimination: 2.13928391801523, difficulty: -1.05258437333796, guessing: 0.0146031435463146 },
    { level: 2, item: "問題", partOfSpeech: "noun", correctAnswer: "issue", distractors: ["pattern", "device", "award"], discrimination: 4.18012123416299, difficulty: -1.20536084608892, guessing: 0.113997447904315 },
    { level: 2, item: "期間", partOfSpeech: "noun", correctAnswer: "term", distractors: ["kid", "platform", "contrast"], discrimination: 2.43637210391788, difficulty: -1.06272799654659, guessing: 0.146862327205805 },
    { level: 2, item: "経済", partOfSpeech: "noun", correctAnswer: "economy", distractors: ["impression", "duty", "cross"], discrimination: 2.57650459700171, difficulty: -2.08806700642603, guessing: 0.0145993822726751 },
    { level: 2, item: "政治の", partOfSpeech: "adjective", correctAnswer: "political", distractors: ["complete", "senior", "sharp"], discrimination: 2.67122773995895, difficulty: -1.76943857013921, guessing: 0.0187028408394886 },
    { level: 2, item: "過程", partOfSpeech: "noun", correctAnswer: "process", distractors: ["sugar", "trick", "channel"], discrimination: 1.84440261351734, difficulty: -1.92395461823059, guessing: 0.0632110947273372 },
    { level: 2, item: "盗む", partOfSpeech: "verb", correctAnswer: "steal", distractors: ["create", "push", "destroy"], discrimination: 4.66843149473234, difficulty: -2.28898599968956, guessing: 0.0624867487841577 },
    { level: 3, item: "種類", partOfSpeech: "noun", correctAnswer: "sort", distractors: ["complaint", "client", "youth"], discrimination: 1.31366061221616, difficulty: -0.947357730797964, guessing: 0.0133729937522661 },
    { level: 3, item: "評議会", partOfSpeech: "noun", correctAnswer: "council", distractors: ["soil", "producer", "strategy"], discrimination: 1.74814492480529, difficulty: -0.755652874349234, guessing: 0.141781655638154 },
    { level: 3, item: "分析", partOfSpeech: "noun", correctAnswer: "analysis", distractors: ["stream", "presence", "monitor"], discrimination: 1.8897480735049, difficulty: -0.56983522901406, guessing: 0.15647490348592 },
    { level: 3, item: "反対", partOfSpeech: "noun", correctAnswer: "opposition", distractors: ["failure", "arrival", "awareness"], discrimination: 1.96290546099529, difficulty: -0.94430920644069, guessing: 0.192074439013158 },
    { level: 3, item: "追う", partOfSpeech: "verb", correctAnswer: "pursue", distractors: ["calculate", "negotiate", "deserve"], discrimination: 2.18222300672536, difficulty: -0.799115631065032, guessing: 0.226647059994051 },
    { level: 3, item: "重要な", partOfSpeech: "adjective", correctAnswer: "significant", distractors: ["immediate", "multiple", "flexible"], discrimination: 1.6858138553401, difficulty: -0.481844616539, guessing: 0.00494630712163104 },
    { level: 3, item: "明らかに", partOfSpeech: "adverb", correctAnswer: "obviously", distractors: ["fairly", "somewhat", "additionally"], discrimination: 1.93026539401, difficulty: -1.12768049121994, guessing: 0.185177938428507 },
    { level: 3, item: "革", partOfSpeech: "noun", correctAnswer: "leather", distractors: ["participant", "reputation", "crack"], discrimination: 1.29125765484797, difficulty: -0.72725841350412, guessing: 0.00929995189302975 },
    { level: 3, item: "比較", partOfSpeech: "noun", correctAnswer: "comparison", distractors: ["quarter", "expertise", "introduction"], discrimination: 2.16583158413454, difficulty: -1.07075949836369, guessing: 0.225466235607567 },
    { level: 3, item: "出会い", partOfSpeech: "noun", correctAnswer: "encounter", distractors: ["stock", "tension", "register"], discrimination: 2.39818797026513, difficulty: -0.65604069154088, guessing: 0.152926161286664 },
    { level: 3, item: "おおよそ", partOfSpeech: "adverb", correctAnswer: "approximately", distractors: ["precisely", "possibly", "slightly"], discrimination: 2.03370700833459, difficulty: -0.135121451630556, guessing: 0.212769991305685 },
    { level: 3, item: "もの（物質）", partOfSpeech: "noun", correctAnswer: "stuff", distractors: ["statistic", "contract", "characteristic"], discrimination: 1.78446681803155, difficulty: -0.130774573413756, guessing: 0.0564139544049375 },
    { level: 3, item: "求める", partOfSpeech: "verb", correctAnswer: "seek", distractors: ["launch", "deny", "abandon"], discrimination: 2.29195860977747, difficulty: -0.23060061967059, guessing: 0.1635439964162 },
    { level: 3, item: "学問の", partOfSpeech: "adjective", correctAnswer: "academic", distractors: ["unlikely", "extraordinary", "mass"], discrimination: 3.58134223678901, difficulty: -0.314660164538795, guessing: 0.222871238563237 },
    { level: 3, item: "視力", partOfSpeech: "noun", correctAnswer: "vision", distractors: ["preparation", "boss", "administration"], discrimination: 1.86158585022274, difficulty: -0.82958340000992, guessing: 0.0317488401788602 },
    { level: 3, item: "欠点", partOfSpeech: "noun", correctAnswer: "fault", distractors: ["reduction", "existence", "description"], discrimination: 1.48683835904175, difficulty: -0.946058065585323, guessing: 0.0461404798934448 },
    { level: 3, item: "身元", partOfSpeech: "noun", correctAnswer: "identity", distractors: ["editor", "wealth", "ancestor"], discrimination: 1.74864779468044, difficulty: -0.995661239150276, guessing: 0.126604955564439 },
    { level: 3, item: "提出する", partOfSpeech: "verb", correctAnswer: "submit", distractors: ["convince", "enable", "satisfy"], discrimination: 1.99277161937758, difficulty: -0.55410267150351, guessing: 0.085305696627542 },
    { level: 3, item: "批判的な", partOfSpeech: "adjective", correctAnswer: "critical", distractors: ["raw", "brief", "unnecessary"], discrimination: 3.344880817409, difficulty: -0.855709542488015, guessing: 0.139687523908515 },
    { level: 3, item: "国内の", partOfSpeech: "adjective", correctAnswer: "domestic", distractors: ["permanent", "constant", "enthusiastic"], discrimination: 1.41135430497581, difficulty: -0.950493983892732, guessing: 0.0103351846809114 },
    { level: 4, item: "信念", partOfSpeech: "noun", correctAnswer: "faith", distractors: ["disposal", "density", "staple"], discrimination: 1.19042656345784, difficulty: -0.11659280194578, guessing: 0.0313076705662978 },
    { level: 4, item: "文字通りに", partOfSpeech: "adverb", correctAnswer: "literally", distractors: ["publicly", "presumably", "formally"], discrimination: 1.56663470317269, difficulty: -0.18312785178973, guessing: 0.0447614851346801 },
    { level: 4, item: "権利を与える", partOfSpeech: "verb", correctAnswer: "entitle", distractors: ["assess", "cultivate", "seize"], discrimination: 4.00038149023987, difficulty: 0.13319565852193, guessing: 0.214696781926316 },
    { level: 4, item: "先の", partOfSpeech: "adjective", correctAnswer: "prior", distractors: ["snowy", "causal", "stressful"], discrimination: 2.25976169332664, difficulty: -0.982839729684233, guessing: 0.151377867440107 },
    { level: 4, item: "大量の", partOfSpeech: "adjective", correctAnswer: "massive", distractors: ["harsh", "primitive", "gradual"], discrimination: 2.25075296473869, difficulty: -0.219098274763004, guessing: 0.18909961894272 },
    { level: 4, item: "角（度）", partOfSpeech: "noun", correctAnswer: "angle", distractors: ["herb", "stove", "manual"], discrimination: 1.40976713569163, difficulty: -0.93348521615256, guessing: 0.011252210576226 },
    { level: 4, item: "歳入", partOfSpeech: "noun", correctAnswer: "revenue", distractors: ["ambition", "coral", "ballet"], discrimination: 4.80131559452022, difficulty: 0.010641752898615, guessing: 0.111191407557468 },
    { level: 4, item: "金融", partOfSpeech: "noun", correctAnswer: "finance", distractors: ["tablet", "farmland", "exploration"], discrimination: 1.90947418569939, difficulty: -0.22671664447743, guessing: 0.206828904881675 },
    { level: 4, item: "知覚する", partOfSpeech: "verb", correctAnswer: "perceive", distractors: ["qualify", "frustrate", "renew"], discrimination: 2.89627298958609, difficulty: -0.267320170204461, guessing: 0.284423315452947 },
    { level: 4, item: "法人の", partOfSpeech: "adjective", correctAnswer: "corporate", distractors: ["occasional", "noticeable", "unavoidable"], discrimination: 1.36076386995634, difficulty: 0.817217664646596, guessing: 0.00681723851716374 },
    { level: 4, item: "統合", partOfSpeech: "noun", correctAnswer: "integration", distractors: ["defeat", "artifact", "wallet"], discrimination: 10.6247925880626, difficulty: -0.165232317538039, guessing: 0.127123723311717 },
    { level: 4, item: "計画", partOfSpeech: "noun", correctAnswer: "scheme", distractors: ["ant", "pamphlet", "merit"], discrimination: 1.74509292414813, difficulty: 0.09466928222817, guessing: 0.221682060662407 },
    { level: 4, item: "収容設備", partOfSpeech: "noun", correctAnswer: "accommodation", distractors: ["miracle", "grace", "conservation"], discrimination: 2.11113247470267, difficulty: -0.0803302669611481, guessing: 0.166036374008193 },
    { level: 4, item: "制限する", partOfSpeech: "verb", correctAnswer: "restrict", distractors: ["depict", "dedicate", "abolish"], discrimination: 5.07142969774894, difficulty: -0.286747550094581, guessing: 0.251970216703725 },
    { level: 4, item: "所有（物）", partOfSpeech: "noun", correctAnswer: "possession", distractors: ["chemist", "personnel", "chamber"], discrimination: 1.20965103078803, difficulty: -0.0638810145340334, guessing: 0.128498365097421 },
    { level: 4, item: "不安", partOfSpeech: "noun", correctAnswer: "anxiety", distractors: ["habitat", "dialect", "convention"], discrimination: 1.64184492285579, difficulty: 0.49638058855936, guessing: 0.0415255037149021 },
    { level: 4, item: "横顔", partOfSpeech: "noun", correctAnswer: "profile", distractors: ["refund", "bloom", "penalty"], discrimination: 1.23771012499516, difficulty: 0.61464838174176, guessing: 0.142330195330679 },
    { level: 4, item: "目に見える", partOfSpeech: "adjective", correctAnswer: "visible", distractors: ["vivid", "pleasant", "stylish"], discrimination: 1.82662278482567, difficulty: -0.50892331846197, guessing: 0.00925848970215405 },
    { level: 4, item: "不動産", partOfSpeech: "noun", correctAnswer: "estate", distractors: ["summit", "mud", "neglect"], discrimination: 3.75780571465459, difficulty: -0.100544164673599, guessing: 0.122787336828416 },
    { level: 4, item: "政治制度", partOfSpeech: "noun", correctAnswer: "regime", distractors: ["voyage", "investigator", "punch"], discrimination: 1.5257039661939, difficulty: 0.467438176490349, guessing: 0.141313454200637 },
    { level: 5, item: "安定", partOfSpeech: "noun", correctAnswer: "stability", distractors: ["astronomer", "relaxation", "dilemma"], discrimination: 1.56985723781376, difficulty: 0.914301963475006, guessing: 0.0262260875003294 },
    { level: 5, item: "選ぶ", partOfSpeech: "verb", correctAnswer: "opt", distractors: ["orient", "deteriorate", "execute"], discrimination: 0.850911101679395, difficulty: 0.139333971940156, guessing: 0.00571038773637942 },
    { level: 5, item: "（〜する）資格のある", partOfSpeech: "adjective", correctAnswer: "eligible", distractors: ["partial", "optimistic", "operational"], discrimination: 2.5221648941349, difficulty: 1.09003300882306, guessing: 0.167774995564354 },
    { level: 5, item: "潜在的に", partOfSpeech: "adverb", correctAnswer: "potentially", distractors: ["scientifically", "remarkably", "adequately"], discrimination: 6.1306060409273, difficulty: 0.958275233210795, guessing: 0.145827140771862 },
    { level: 5, item: "格付け", partOfSpeech: "noun", correctAnswer: "rating", distractors: ["homeland", "discomfort", "logo"], discrimination: 1.26175626938899, difficulty: 1.18169539271953, guessing: 0.0118297715201111 },
    { level: 5, item: "敵意のある", partOfSpeech: "adjective", correctAnswer: "hostile", distractors: ["advisory", "mandatory", "magnetic"], discrimination: 2.08343811855697, difficulty: 0.122256006463701, guessing: 0.00451649257773473 },
    { level: 5, item: "残酷な", partOfSpeech: "adjective", correctAnswer: "cruel", distractors: ["serial", "exotic", "unrelated"], discrimination: 1.96416346875482, difficulty: 1.33863242499781, guessing: 0.117735111403888 },
    { level: 5, item: "身をさらすこと", partOfSpeech: "noun", correctAnswer: "exposure", distractors: ["supplement", "hydrogen", "detention"], discrimination: 2.21298540719469, difficulty: 0.59827152701287, guessing: 0.184458901965183 },
    { level: 5, item: "降下", partOfSpeech: "noun", correctAnswer: "descent", distractors: ["frontier", "peer", "diesel"], discrimination: 4.26528842252328, difficulty: -0.0523743956700169, guessing: 0.16994823386578 },
    { level: 5, item: "一覧表への記入", partOfSpeech: "noun", correctAnswer: "listing", distractors: ["tolerance", "protocol", "ramp"], discrimination: 0.998495827260258, difficulty: 0.666390082140726, guessing: 0.0127496015014125 },
    { level: 5, item: "協力", partOfSpeech: "noun", correctAnswer: "collaboration", distractors: ["ambassador", "creativity", "lawsuit"], discrimination: 2.01547939330527, difficulty: 0.09014807289226, guessing: 0.0866785601403839 },
    { level: 5, item: "業績", partOfSpeech: "noun", correctAnswer: "accomplishment", distractors: ["constraint", "resignation", "inventory"], discrimination: 1.97944685099946, difficulty: -0.319507591077966, guessing: 0.204541613510672 },
    { level: 5, item: "使用料", partOfSpeech: "noun", correctAnswer: "toll", distractors: ["sickness", "pharmacy", "patch"], discrimination: 1.74276385089103, difficulty: 1.12089628973942, guessing: 0.146034923781463 },
    { level: 5, item: "枠組み", partOfSpeech: "noun", correctAnswer: "framework", distractors: ["penny", "revision", "rejection"], discrimination: 1.43209977676786, difficulty: 1.49814656320327, guessing: 0.010330823651947 },
    { level: 5, item: "墓", partOfSpeech: "noun", correctAnswer: "tomb", distractors: ["inclusion", "separation", "surplus"], discrimination: 1.63358042873748, difficulty: 1.52961077514597, guessing: 0.0382766244803406 },
    { level: 5, item: "祈り", partOfSpeech: "noun", correctAnswer: "prayer", distractors: ["globe", "nitrogen", "ram"], discrimination: 1.76344534398845, difficulty: 0.867110490836613, guessing: 0.14676068051031 },
    { level: 5, item: "一致", partOfSpeech: "noun", correctAnswer: "correspondence", distractors: ["worship", "transplant", "combat"], discrimination: 2.10175210842634, difficulty: 0.470788272418638, guessing: 0.228310848751147 },
    { level: 5, item: "直立した", partOfSpeech: "verb", correctAnswer: "erect", distractors: ["tremble", "deprive", "renovate"], discrimination: 0.897421945862239, difficulty: 0.732139681515033, guessing: 0.0236055336147904 },
    { level: 5, item: "完全な", partOfSpeech: "adjective", correctAnswer: "integral", distractors: ["offshore", "legislative", "triple"], discrimination: 4.66451819463894, difficulty: 1.14920427244563, guessing: 0.189474550863229 },
    { level: 5, item: "抽象的な", partOfSpeech: "adjective", correctAnswer: "abstract", distractors: ["inherent", "naive", "archaeological"], discrimination: 1.58051519448624, difficulty: 0.872775379588432, guessing: 0.0153602855887117 },
    { level: 6, item: "とても天気の良い", partOfSpeech: "adjective", correctAnswer: "gorgeous", distractors: ["lightweight", "inferior", "promotional"], discrimination: 1.02509601292516, difficulty: 1.74977139766063, guessing: 0.10717521203576 },
    { level: 6, item: "それゆえ", partOfSpeech: "adverb", correctAnswer: "hence", distractors: ["urgently", "ecologically", "nicely"], discrimination: 3.26656583114405, difficulty: 0.139063365775562, guessing: 0.14077339160043 },
    { level: 6, item: "概要", partOfSpeech: "noun", correctAnswer: "overview", distractors: ["spacecraft", "alteration", "mound"], discrimination: 1.17156819532336, difficulty: 0.38524165472821, guessing: 0.0145116723027747 },
    { level: 6, item: "祖先", partOfSpeech: "noun", correctAnswer: "ancestry", distractors: ["coordination", "liability", "mathematician"], discrimination: 5.40607580227347, difficulty: 1.84934475908173, guessing: 0.223867422501257 },
    { level: 6, item: "子孫", partOfSpeech: "noun", correctAnswer: "descendant", distractors: ["capitalism", "acquaintance", "ankle"], discrimination: 2.18665198422626, difficulty: 1.47635607630482, guessing: 0.217120251517674 },
    { level: 6, item: "実験", partOfSpeech: "noun", correctAnswer: "experimentation", distractors: ["sunglass", "courtesy", "pathogen"], discrimination: 1.9508835701201, difficulty: 0.099572258518679, guessing: 0.00999937499014052 },
    { level: 6, item: "原子", partOfSpeech: "noun", correctAnswer: "atom", distractors: ["fungus", "rust", "fist"], discrimination: 2.14266931990076, difficulty: 1.91868636867566, guessing: 0.0107989473549911 },
    { level: 6, item: "傾向", partOfSpeech: "noun", correctAnswer: "inclination", distractors: ["projector", "resin", "infancy"], discrimination: 1.18822414752333, difficulty: 1.15350193786362, guessing: 0.0124851646162233 },
    { level: 6, item: "自由にする", partOfSpeech: "verb", correctAnswer: "liberate", distractors: ["encompass", "alleviate", "circulate"], discrimination: 2.24143919420591, difficulty: 1.57286562471018, guessing: 0.140980941610716 },
    { level: 6, item: "誠実な", partOfSpeech: "adjective", correctAnswer: "sincere", distractors: ["parliamentary", "linear", "inward"], discrimination: 2.2602133416099, difficulty: 0.23138273995055, guessing: 0.237014799407696 },
    { level: 6, item: "市民の", partOfSpeech: "adjective", correctAnswer: "civilian", distractors: ["oval", "simultaneous", "damp"], discrimination: 2.45954285715963, difficulty: 2.10402717994352, guessing: 0.0230752311391577 },
    { level: 6, item: "統計の", partOfSpeech: "adjective", correctAnswer: "statistical", distractors: ["radioactive", "unpopular", "jumbo"], discrimination: 1.32850232611611, difficulty: 1.51988134415988, guessing: 0.0305211503392946 },
    { level: 6, item: "居住", partOfSpeech: "noun", correctAnswer: "residency", distractors: ["commander", "aerial", "lieutenant"], discrimination: 1.60405707450354, difficulty: 0.625818207266, guessing: 0.00978478541798415 },
    { level: 6, item: "不一致", partOfSpeech: "noun", correctAnswer: "discrepancy", distractors: ["incorporation", "patron", "breadth"], discrimination: 1.76895119202939, difficulty: 1.25963680077614, guessing: 0.0162884373815863 },
    { level: 6, item: "加入する", partOfSpeech: "verb", correctAnswer: "subscribe", distractors: ["withhold", "differentiate", "murmur"], discrimination: 1.50462545884541, difficulty: 1.05191973517141, guessing: 0.183358243975552 },
    { level: 6, item: "抑制する", partOfSpeech: "verb", correctAnswer: "restrain", distractors: ["initiate", "comprehend", "intrigue"], discrimination: 1.71892199423027, difficulty: 1.50892509528772, guessing: 0.0663753583044939 },
    { level: 6, item: "従う", partOfSpeech: "verb", correctAnswer: "conform", distractors: ["recite", "evacuate", "compel"], discrimination: 1.49916831554065, difficulty: 1.70422018849171, guessing: 0.0450243434175983 },
    { level: 6, item: "嘆き", partOfSpeech: "noun", correctAnswer: "grief", distractors: ["rifle", "miner", "naturalist"], discrimination: 1.19065446201934, difficulty: 1.46568778219249, guessing: 0.0674641759392268 },
    { level: 6, item: "確認", partOfSpeech: "noun", correctAnswer: "verification", distractors: ["stagnation", "deduction", "herd"], discrimination: 1.35707650831689, difficulty: 0.0632854671962239, guessing: 0.0276682657073602 },
    { level: 6, item: "分割", partOfSpeech: "noun", correctAnswer: "partition", distractors: ["motto", "pickle", "novelty"], discrimination: 1.43654792399002, difficulty: 1.15680514398148, guessing: 0.0753880717873424 },
    { level: 7, item: "誓い", partOfSpeech: "noun", correctAnswer: "oath", distractors: ["repertoire", "payroll", "ensemble"], discrimination: 2.60741836569888, difficulty: 1.25399062767499, guessing: 0.152952790053737 },
    { level: 7, item: "保証書", partOfSpeech: "noun", correctAnswer: "warrant", distractors: ["unison", "aerospace", "gospel"], discrimination: 1.17828362889223, difficulty: 1.07949720364868, guessing: 0.144935933314544 },
    { level: 7, item: "貨物", partOfSpeech: "noun", correctAnswer: "freight", distractors: ["ignorance", "mucosa", "mainland"], discrimination: 1.58269381162658, difficulty: 2.49392744698618, guessing: 0.125434293697562 },
    { level: 7, item: "主張", partOfSpeech: "noun", correctAnswer: "assertion", distractors: ["inquest", "biopsy", "disclosure"], discrimination: 1.52638505905812, difficulty: 1.12808487708165, guessing: 0.179224322942926 },
    { level: 7, item: "固定観念", partOfSpeech: "noun", correctAnswer: "stereotype", distractors: ["precedent", "module", "dismay"], discrimination: 1.81549047464789, difficulty: 0.716500310749043, guessing: 0.173795565829433 },
    { level: 7, item: "補足の ", partOfSpeech: "adjective", correctAnswer: "complementary", distractors: ["imminent", "tricky", "eternal"], discrimination: 1.95339535482388, difficulty: 1.33081371717151, guessing: 0.230538482520176 },
    { level: 7, item: "運転する人", partOfSpeech: "noun", correctAnswer: "motorist", distractors: ["secrecy", "curator", "analogy"], discrimination: 1.00988179439087, difficulty: 0.359536398611924, guessing: 0.0176112251307963 },
    { level: 7, item: "ちらりと見えること", partOfSpeech: "noun", correctAnswer: "glimpse", distractors: ["betrayal", "disco", "banner"], discrimination: 8.23597120882976, difficulty: 1.95796244789731, guessing: 0.245054472766734 },
    { level: 7, item: "侵入者", partOfSpeech: "noun", correctAnswer: "raider", distractors: ["monopoly", "skipper", "cinnamon"], discrimination: 0.910529152971927, difficulty: 1.10984671699923, guessing: 0.021885331625089 },
    { level: 7, item: "外交", partOfSpeech: "noun", correctAnswer: "diplomacy", distractors: ["injunction", "farmhouse", "embargo"], discrimination: 1.70469268649998, difficulty: 1.12189814826731, guessing: 0.180829135717648 },
    { level: 7, item: "バランスをとる", partOfSpeech: "verb", correctAnswer: "poise", distractors: ["chuck", "knit", "shiver"], discrimination: 1.6441705717507, difficulty: 2.07242907105954, guessing: 0.0122257556432279 },
    { level: 7, item: "装飾用の ", partOfSpeech: "adjective", correctAnswer: "decorative", distractors: ["aging", "hopeless", "marital"], discrimination: 1.85038624528662, difficulty: 1.42491571401915, guessing: 0.0293362119739844 },
    { level: 7, item: "心臓の ", partOfSpeech: "adjective", correctAnswer: "cardiac", distractors: ["creamy", "apt", "dire"], discrimination: 1.23802612907763, difficulty: 2.12893900162273, guessing: 0.0587600192507367 },
    { level: 7, item: "量の ", partOfSpeech: "adjective", correctAnswer: "quantitative", distractors: ["relentless", "median", "comparative"], discrimination: 1.70504202511163, difficulty: 1.3074718985076, guessing: 0.0139604071931999 },
    { level: 7, item: "深く ", partOfSpeech: "adverb", correctAnswer: "profoundly", distractors: ["loosely", "indirectly", "etcetera"], discrimination: 3.36229930554489, difficulty: 2.02110227312478, guessing: 0.262957041836544 },
    { level: 7, item: "選挙区", partOfSpeech: "noun", correctAnswer: "constituency", distractors: ["unrest", "majesty", "baron"], discrimination: 1.42926181704341, difficulty: 1.76268024310698, guessing: 0.067466367939401 },
    { level: 7, item: "格言", partOfSpeech: "noun", correctAnswer: "maxim", distractors: ["brigade", "accountant", "probation"], discrimination: 4.92180781847632, difficulty: 1.96546290062054, guessing: 0.0746326286748535 },
    { level: 7, item: "軽蔑", partOfSpeech: "noun", correctAnswer: "contempt", distractors: ["communism", "stimulation", "referendum"], discrimination: 2.44854227491467, difficulty: 1.89076986009438, guessing: 0.195611865831798 },
    { level: 7, item: "頑固な", partOfSpeech: "adjective", correctAnswer: "stubborn", distractors: ["pragmatic", "bureaucratic", "molecular"], discrimination: 2.25791775780656, difficulty: 1.45342368361068, guessing: 0.270928236561614 },
    { level: 7, item: "屋根裏部屋 ", partOfSpeech: "noun", correctAnswer: "attic", distractors: ["referral", "embargo", "goodwill"], discrimination: 0.918663065920893, difficulty: 1.49304265434803, guessing: 0.0103438512439508 },
    { level: 8, item: "構文", partOfSpeech: "noun", correctAnswer: "syntax", distractors: ["brewery", "hooker", "hawk"], discrimination: 1.07360255848345, difficulty: 2.31349868490648, guessing: 0.00748289168415301 },
    { level: 8, item: "起訴する", partOfSpeech: "verb", correctAnswer: "prosecute", distractors: ["grate", "lurk", "rinse"], discrimination: 1.96943695121481, difficulty: 2.0674799673427, guessing: 0.258245338874227 },
    { level: 8, item: "普遍的に", partOfSpeech: "adverb", correctAnswer: "universally", distractors: ["inadvertently", "boldly", "immensely"], discrimination: 4.64481934954136, difficulty: 1.0399844016976, guessing: 0.397516784895916 },
    { level: 8, item: "匿名であること", partOfSpeech: "noun", correctAnswer: "anonymity", distractors: ["broom", "snooker", "saloon"], discrimination: 1.15138024281073, difficulty: 2.23275149723823, guessing: 0.0102320433851838 },
    { level: 8, item: "子宮", partOfSpeech: "noun", correctAnswer: "womb", distractors: ["envy", "vibration", "cooker"], discrimination: 0.860980294955352, difficulty: 2.92369605073857, guessing: 0.0396921561653864 },
    { level: 8, item: "正気の", partOfSpeech: "adjective", correctAnswer: "sane", distractors: ["penal", "pristine", "pancreatic"], discrimination: 3.39545191078376, difficulty: 2.95679576785815, guessing: 0.170552335138351 },
    { level: 8, item: "自殺の", partOfSpeech: "adjective", correctAnswer: "suicidal", distractors: ["longitudinal", "disruptive", "endoscopic"], discrimination: 1.87002139881424, difficulty: 0.872428416497446, guessing: 0.05104361688409 },
    { level: 8, item: "聖歌", partOfSpeech: "noun", correctAnswer: "anthem", distractors: ["tempo", "elegance", "slaughter"], discrimination: 0.902904670088228, difficulty: 2.03083961606355, guessing: 0.0279555750397444 },
    { level: 8, item: "説得", partOfSpeech: "noun", correctAnswer: "persuasion", distractors: ["persona", "variability", "heartbeat"], discrimination: 2.3948441632652, difficulty: 1.75418166905498, guessing: 0.21614043119561 },
    { level: 8, item: "長期間続く", partOfSpeech: "adjective", correctAnswer: "perennial", distractors: ["homogeneous", "psychic", "fluorescent"], discrimination: 2.01767898127296, difficulty: 2.52353202227224, guessing: 0.203784538651509 },
    { level: 8, item: "結果として起こる", partOfSpeech: "adjective", correctAnswer: "consequent", distractors: ["doctoral", "stray", "recurrent"], discrimination: 2.05876976555754, difficulty: 1.48295031200976, guessing: 0.226345831714249 },
    { level: 8, item: "没頭", partOfSpeech: "noun", correctAnswer: "preoccupation", distractors: ["monsieur", "starvation", "gypsy"], discrimination: 3.64431281339349, difficulty: 1.06310995883635, guessing: 0.26287661781997 },
    { level: 8, item: "恐怖", partOfSpeech: "noun", correctAnswer: "dread", distractors: ["relegation", "deviation", "midfielder"], discrimination: 3.15922855113942, difficulty: 1.57408899530888, guessing: 0.260122748329134 },
    { level: 8, item: "法令", partOfSpeech: "noun", correctAnswer: "decree", distractors: ["affinity", "devastation", "parasite"], discrimination: 2.72489143939076, difficulty: 2.67190277547001, guessing: 0.12506128887051 },
    { level: 8, item: "贅沢な", partOfSpeech: "adjective", correctAnswer: "luxurious", distractors: ["erratic", "cunning", "indicative"], discrimination: 2.48315481731094, difficulty: 0.56651287033554, guessing: 0.0840493847881545 },
    { level: 8, item: "省略", partOfSpeech: "noun", correctAnswer: "omission", distractors: ["gem", "resonance", "retaliation"], discrimination: 0.887333068962794, difficulty: 2.70470401170826, guessing: 0.0442406502868085 },
    { level: 8, item: "移り変わり", partOfSpeech: "noun", correctAnswer: "flux", distractors: ["coating", "whereabouts", "jeopardy"], discrimination: 3.61520963360786, difficulty: 3.08824798344858, guessing: 0.238070649066161 },
    { level: 8, item: "説教", partOfSpeech: "noun", correctAnswer: "sermon", distractors: ["apprehension", "brilliance", "pilgrimage"], discrimination: 1.6758466387719, difficulty: 2.65247446326893, guessing: 0.100311119562815 },
    { level: 8, item: "雑用", partOfSpeech: "noun", correctAnswer: "chore", distractors: ["wig", "fracture", "foreground"], discrimination: 1.99523765223184, difficulty: 2.46088318485034, guessing: 0.176671951705064 },
    { level: 8, item: "階層性の", partOfSpeech: "adjective", correctAnswer: "hierarchical", distractors: ["aristocratic", "martial", "proprietary"], discrimination: 1.66713500928888, difficulty: 2.7998834555433, guessing: 0.052274658477023 }
];

// POSの日本語表示用マッピング
const posMapping = {
    "noun": "名詞",
    "verb": "動詞",
    "adjective": "形容詞",
    "adverb": "副詞"
};

// グローバル変数
let currentQuestionIndex = 0;
let currentQuestion = null;
let results = {
    responses: [],
    correctCount: 0,
    levelScores: {},
    startTime: null,
    reactionTimes: []
};

// DOM要素
const startButton = document.getElementById('startButton');
const startScreen = document.getElementById('startScreen');
const testContainer = document.getElementById('testContainer');
const fixationPoint = document.getElementById('fixationPoint');
const questionArea = document.getElementById('questionArea');
const currentQuestionElem = document.getElementById('currentQuestion');
const optionsContainer = document.getElementById('optionsContainer');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const currentLevelElem = document.getElementById('currentLevel');
const currentPOSElem = document.getElementById('currentPOS');
const resultsDiv = document.getElementById('results');
const correctAnswersElem = document.getElementById('correctAnswers');
const totalQuestionsElem = document.getElementById('totalQuestions');
const vocabularySizeElem = document.getElementById('vocabularySize');
const averageRTElem = document.getElementById('averageRT');
const levelBreakdownElem = document.getElementById('levelBreakdown');
const restartButton = document.getElementById('restartButton');

// イベントリスナー
startButton.addEventListener('click', startTest);
restartButton.addEventListener('click', () => {
    resultsDiv.style.display = 'none';
    startScreen.style.display = 'block';
});
document.addEventListener('keydown', handleKeyboardInput);

// 印刷ボタン
document.getElementById('printButton').addEventListener('click', () => {
    window.print();
});

// データダウンロードボタン（Excel形式）
document.getElementById('downloadExcelButton').addEventListener('click', () => downloadTestData('excel'));

// データダウンロードボタン（JSON形式）
document.getElementById('downloadJsonButton').addEventListener('click', () => downloadTestData('json'));

// レベル間の休憩から次へ進むボタン
document.getElementById('continueButton').addEventListener('click', continueToNextLevel);

// レベル間の休憩からテストを継続する関数
function continueToNextLevel() {
    // 休憩画面を非表示にし、テスト画面を表示
    document.getElementById('breakScreen').style.display = 'none';
    testContainer.style.display = 'block';
    
    // 注視点を表示して次の問題に進む準備
    fixationPoint.textContent = '+';
    fixationPoint.style.display = 'block';
    questionArea.style.display = 'none';
    
    // 一定時間後に次の問題を読み込む
    setTimeout(() => {
        loadQuestion();
    }, 1000);
}

// テスト開始関数
function startTest() {
    startScreen.style.display = 'none';
    testContainer.style.display = 'block';
    results = {
        responses: [],
        correctCount: 0,
        levelScores: {},
        startTime: new Date(),
        reactionTimes: []
    };
    currentQuestionIndex = 0;
    
    // カウントダウンアニメーション
    fixationPoint.textContent = '3';
    fixationPoint.style.display = 'block';
    questionArea.style.display = 'none';
    
    setTimeout(() => {
        fixationPoint.textContent = '2';
        setTimeout(() => {
            fixationPoint.textContent = '1';
            setTimeout(() => {
                fixationPoint.textContent = 'スタート!';
                setTimeout(() => {
                    fixationPoint.textContent = '+';
                    loadQuestion();
                }, 800);
            }, 800);
        }, 800);
    }, 800);
}

// テストデータのダウンロード関数
function downloadTestData(format = 'json') {
    // 基本データを準備
    const summaryData = {
        testDate: results.startTime,
        totalQuestions: testData.length,
        correctAnswers: results.correctCount,
        estimatedVocabularySize: estimateVocabularySize(),
        averageReactionTime: results.reactionTimes.reduce((sum, time) => sum + time, 0) / results.reactionTimes.length,
        abilityEstimate: estimateAbility()
    };
    
    // 詳細な反応データを準備
    const responseData = results.responses.map(r => {
        const item = testData[r.questionIndex];
        return {
            level: item.level,
            item: item.item,
            partOfSpeech: item.partOfSpeech,
            correctAnswer: item.correctAnswer,
            selectedAnswer: r.isDontKnow ? "わからない" : (r.isCorrect ? item.correctAnswer : [item.correctAnswer, ...item.distractors][r.selectedIndex]),
            isCorrect: r.isCorrect,
            isDontKnow: r.isDontKnow,
            reactionTime: r.reactionTime,
            discrimination: item.discrimination,
            difficulty: item.difficulty,
            guessing: item.guessing
        };
    });
    
    // レベル別の集計データを準備
    const levelData = [];
    for (const level in results.levelScores) {
        const score = results.levelScores[level];
        // このレベルの回答だけをフィルタリング
        const levelResponses = results.responses.filter(r => testData[r.questionIndex].level == level);
        // 平均反応時間を計算
        const avgRT = levelResponses.reduce((sum, r) => sum + r.reactionTime, 0) / levelResponses.length;
        
        levelData.push({
            level: parseInt(level),
            total: score.total,
            correct: score.correct,
            correctRate: score.correct / score.total,
            avgReactionTime: avgRT
        });
    }
    
    if (format === 'json') {
        // JSONとしてダウンロード
        const exportData = {
            summary: summaryData,
            levelResults: levelData,
            responses: responseData
        };
        
        const dataStr = JSON.stringify(exportData, null, 2);
        const dataBlob = new Blob([dataStr], {type: 'application/json'});
        
        const downloadLink = document.createElement('a');
        downloadLink.href = URL.createObjectURL(dataBlob);
        downloadLink.download = `vst-nj8-results-${new Date().toISOString().slice(0, 10)}.json`;
        
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
    } 
    else if (format === 'excel') {
        try {
            // Excel形式でダウンロード（SheetJSを使用）
            // ワークブックを作成
            const wb = XLSX.utils.book_new();
            
            // シート1: 概要情報
            const summarySheet = [
                ['VST-NJ8 テスト結果概要'],
                [''],
                ['テスト日時', summaryData.testDate.toLocaleString()],
                ['問題数', summaryData.totalQuestions],
                ['正答数', summaryData.correctAnswers],
                ['正答率', `${(summaryData.correctAnswers / summaryData.totalQuestions * 100).toFixed(1)}%`],
                ['推定語彙サイズ', summaryData.estimatedVocabularySize],
                ['平均反応時間', `${summaryData.averageReactionTime.toFixed(2)}秒`],
                ['能力値推定（θ）', summaryData.abilityEstimate.toFixed(3)]
            ];
            
            const ws1 = XLSX.utils.aoa_to_sheet(summarySheet);
            XLSX.utils.book_append_sheet(wb, ws1, '概要');
            
            // セルの書式設定
            if (!ws1['!cols']) ws1['!cols'] = [];
            ws1['!cols'][0] = { wch: 15 }; // 列A幅
            ws1['!cols'][1] = { wch: 20 }; // 列B幅
            
            // シート2: レベル別結果
            const levelSheetHeader = [
                ['レベル', '問題数', '正答数', '正答率', '平均反応時間（秒）']
            ];
            
            const levelSheetData = levelData.sort((a, b) => a.level - b.level).map(level => [
                level.level,
                level.total,
                level.correct,
                level.correctRate.toFixed(3),
                level.avgReactionTime.toFixed(2)
            ]);
            
            const ws2 = XLSX.utils.aoa_to_sheet([...levelSheetHeader, ...levelSheetData]);
            XLSX.utils.book_append_sheet(wb, ws2, 'レベル別結果');
            
            // シート3: 詳細な回答データ
            const responseSheetHeader = [
                ['レベル', '日本語', '品詞', '正解', '選択した回答', '正誤', '「わからない」選択', '反応時間（秒）', '識別力', '困難度', '推測パラメータ']
            ];
            
            const responseSheetData = responseData.map(r => [
                r.level,
                r.item,
                r.partOfSpeech,
                r.correctAnswer,
                r.selectedAnswer,
                r.isCorrect ? '○' : '×',
                r.isDontKnow ? '○' : '×',
                r.reactionTime.toFixed(2),
                r.discrimination.toFixed(3),
                r.difficulty.toFixed(3),
                r.guessing.toFixed(3)
            ]);
            
            const ws3 = XLSX.utils.aoa_to_sheet([...responseSheetHeader, ...responseSheetData]);
            XLSX.utils.book_append_sheet(wb, ws3, '詳細データ');
            
            // ファイル名を設定してダウンロード
            const filename = `vst-nj8-results-${new Date().toISOString().slice(0, 10)}.xlsx`;
            XLSX.writeFile(wb, filename);
        } catch (error) {
            console.error("Excel出力中にエラーが発生しました:", error);
            alert("Excel出力中にエラーが発生しました。JSONフォーマットを試してください。");
        }
    }
}

// 問題を読み込む関数
function loadQuestion() {
    if (currentQuestionIndex >= testData.length) {
        finishTest();
        return;
    }

    // 進捗状況の更新
    updateProgress();
    
    // 注視点を表示
    fixationPoint.style.display = 'block';
    questionArea.style.display = 'none';
    
    // 一定時間後に問題を表示
    setTimeout(() => {
        fixationPoint.style.display = 'none';
        questionArea.style.display = 'block';
        
        // 問題データを取得
        currentQuestion = testData[currentQuestionIndex];
        
        // 問題表示を更新
        currentQuestionElem.textContent = currentQuestion.item;
        currentLevelElem.textContent = currentQuestion.level;
        currentPOSElem.textContent = posMapping[currentQuestion.partOfSpeech] || currentQuestion.partOfSpeech;
        
        // 選択肢を表示
        displayOptions(currentQuestion);
        
        // 反応時間の計測開始
        results.responseStartTime = new Date();
    }, 1000); // 1秒間注視点を表示
}

// 選択肢表示関数
function displayOptions(question) {
    // 選択肢コンテナをクリア
    optionsContainer.innerHTML = '';
    
    // 正解と不正解をまとめる
    let options = [question.correctAnswer, ...question.distractors];
    
    // 選択肢をシャッフル（ランダム化）
    options = shuffleArray(options);
    
    // 正解の位置を記録
    const correctIndex = options.indexOf(question.correctAnswer);
    
    // 選択肢を表示
    options.forEach((option, index) => {
        const optionDiv = document.createElement('div');
        optionDiv.className = 'option';
        optionDiv.innerHTML = `<span class="option-key">${index + 1}</span> ${option}`;
        optionDiv.dataset.index = index;
        optionDiv.dataset.value = option;
        optionDiv.addEventListener('click', () => handleAnswer(index));
        optionsContainer.appendChild(optionDiv);
    });
    
    // 「わからない」オプションを追加
    const dontKnowDiv = document.createElement('div');
    dontKnowDiv.className = 'option';
    dontKnowDiv.innerHTML = `<span class="option-key">5</span> わからない`;
    dontKnowDiv.dataset.index = 4;
    dontKnowDiv.dataset.value = 'dontknow';
    dontKnowDiv.addEventListener('click', () => handleAnswer(4));
    optionsContainer.appendChild(dontKnowDiv);
    
    // 正解の位置を更新
    currentQuestion.correctIndex = correctIndex;
}

// キーボード入力処理関数
function handleKeyboardInput(event) {
    if (!questionArea.style.display || questionArea.style.display === 'none') return;
    
    const key = parseInt(event.key);
    if (key >= 1 && key <= 5) {
        handleAnswer(key - 1);
    }
}

// 回答処理関数
function handleAnswer(selectedIndex) {
    // 反応時間を計算
    const endTime = new Date();
    const reactionTime = (endTime - results.responseStartTime) / 1000; // 秒単位
    
    // 正解判定
    const isCorrect = selectedIndex === currentQuestion.correctIndex;
    // 「わからない」を選択した場合は不正解
    const isDontKnow = selectedIndex === 4;
    
    // 結果を保存
    results.responses.push({
        questionIndex: currentQuestionIndex,
        question: currentQuestion.item,
        level: currentQuestion.level,
        selectedIndex: selectedIndex,
        correctIndex: currentQuestion.correctIndex,
        isCorrect: isCorrect,
        isDontKnow: isDontKnow,
        reactionTime: reactionTime
    });
    
    // 正解数をカウント
    if (isCorrect) results.correctCount++;
    
    // レベル別スコアを記録
    if (!results.levelScores[currentQuestion.level]) {
        results.levelScores[currentQuestion.level] = {
            total: 0,
            correct: 0
        };
    }
    results.levelScores[currentQuestion.level].total++;
    if (isCorrect) results.levelScores[currentQuestion.level].correct++;
    
    // 反応時間を記録
    results.reactionTimes.push(reactionTime);
    
    // レベル間の休憩をチェック
    const nextQuestionIndex = currentQuestionIndex + 1;
    if (nextQuestionIndex < testData.length) {
        const currentLevel = currentQuestion.level;
        const nextLevel = testData[nextQuestionIndex].level;
        
        if (currentLevel !== nextLevel) {
            // レベルが変わる場合は休憩画面を表示
            showBreakScreen(currentLevel, nextQuestionIndex);
            return;
        }
    }
    
    // 次の問題へ
    currentQuestionIndex++;
    setTimeout(loadQuestion, 500);
}

// レベル間の休憩画面を表示する関数
function showBreakScreen(completedLevel, nextQuestionIndex) {
    testContainer.style.display = 'none';
    
    // 休憩画面の情報を更新
    document.getElementById('completedLevel').textContent = completedLevel;
    document.getElementById('breakProgress').textContent = nextQuestionIndex;
    
    // レベルの結果サマリーを作成
    const levelSummary = document.getElementById('levelSummary');
    const levelScore = results.levelScores[completedLevel];
    const correctRate = (levelScore.correct / levelScore.total * 100).toFixed(1);
    
    levelSummary.innerHTML = `
        <h3>レベル${completedLevel}の結果</h3>
        <p>正答数: ${levelScore.correct} / ${levelScore.total}</p>
        <p>正答率: ${correctRate}%</p>
    `;
    
    // 休憩画面を表示
    document.getElementById('breakScreen').style.display = 'block';
    
    // 次の問題インデックスを更新
    currentQuestionIndex = nextQuestionIndex;
}

// 進捗状況更新関数
function updateProgress() {
    const progress = (currentQuestionIndex / testData.length) * 100;
    progressFill.style.width = `${progress}%`;
    
    // 現在のレベルを取得
    const currentLevel = testData[currentQuestionIndex].level;
    
    // 現在のレベル内の問題番号を計算
    const levelStartIndex = testData.findIndex(item => item.level === currentLevel);
    const levelQuestionNumber = currentQuestionIndex - levelStartIndex + 1;
    
    // 進捗テキストを更新（全体の進捗とレベル内の進捗を表示）
    progressText.textContent = `レベル ${currentLevel} - 問題 ${levelQuestionNumber}/20 (全体: ${currentQuestionIndex + 1}/${testData.length})`;
}

// テスト終了関数
function finishTest() {
    testContainer.style.display = 'none';
    resultsDiv.style.display = 'block';
    
    // 結果を表示
    correctAnswersElem.textContent = results.correctCount;
    totalQuestionsElem.textContent = testData.length;
    
    // 平均反応時間を計算
    const avgRT = results.reactionTimes.reduce((sum, time) => sum + time, 0) / results.reactionTimes.length;
    averageRTElem.textContent = avgRT.toFixed(2);
    
    // 語彙サイズの推定（IRT分析に基づく）
    const estimatedSize = estimateVocabularySize();
    vocabularySizeElem.textContent = estimatedSize.toLocaleString();
    
    // 「わからない」選択の割合を計算
    const dontKnowCount = results.responses.filter(r => r.isDontKnow).length;
    const dontKnowPercentage = ((dontKnowCount / results.responses.length) * 100).toFixed(1);
    
    // 推測率（guessingパラメータの活用）
    const totalGuessing = testData.reduce((sum, item) => sum + item.guessing, 0);
    const averageGuessing = (totalGuessing / testData.length * 100).toFixed(1);
    
    // レベル別の結果を表示
    displayLevelBreakdown(dontKnowPercentage, averageGuessing);
}

// レベル別結果表示関数
function displayLevelBreakdown(dontKnowPercentage, averageGuessing) {
    levelBreakdownElem.innerHTML = '<h3>テスト結果分析</h3>';
    
    // 基本情報の表示
    let breakdownHTML = `
        <div style="margin-bottom: 20px;">
            <p>IRTモデルによる能力推定値（θ）: <span id="abilityValue">0.00</span></p>
            <p>「わからない」と回答した割合: ${dontKnowPercentage}%</p>
            <p>テスト項目の平均推測パラメータ: ${averageGuessing}%</p>
        </div>`;
    
    // テーブル形式のレベル別結果
    breakdownHTML += '<table style="width:100%; border-collapse: collapse; margin-top: 15px;">';
    breakdownHTML += '<tr><th style="border: 1px solid #ddd; padding: 8px; text-align: left;">レベル</th><th style="border: 1px solid #ddd; padding: 8px; text-align: left;">正答率</th><th style="border: 1px solid #ddd; padding: 8px; text-align: left;">平均反応時間</th></tr>';
    
    // レベル別の平均反応時間を計算
    const levelRTs = {};
    results.responses.forEach(response => {
        const level = testData[response.questionIndex].level;
        if (!levelRTs[level]) {
            levelRTs[level] = {
                total: 0,
                count: 0
            };
        }
        levelRTs[level].total += response.reactionTime;
        levelRTs[level].count++;
    });
    
    for (const level in results.levelScores) {
        const score = results.levelScores[level];
        const percentage = (score.correct / score.total * 100).toFixed(1);
        
        // 平均反応時間
        const avgRT = levelRTs[level] ? (levelRTs[level].total / levelRTs[level].count).toFixed(2) : "N/A";
        
        breakdownHTML += `<tr>
            <td style="border: 1px solid #ddd; padding: 8px;">レベル ${level}</td>
            <td style="border: 1px solid #ddd; padding: 8px;">${percentage}% (${score.correct}/${score.total})</td>
            <td style="border: 1px solid #ddd; padding: 8px;">${avgRT} 秒</td>
        </tr>`;
    }
    
    breakdownHTML += '</table>';
    
    // グラフの追加
    breakdownHTML += createAbilityChart();
    breakdownHTML += createReactionTimeChart();
    
    levelBreakdownElem.innerHTML += breakdownHTML;
    
    // 能力値の推定と表示
    const theta = estimateAbility();
    document.getElementById('abilityValue').textContent = theta.toFixed(2);
    
    // 詳細な分析結果を表示（オプション）
    displayDetailedAnalysis();
}

// 詳細な分析結果表示関数
function displayDetailedAnalysis() {
    // わからない選択と反応時間の関係などの詳細分析
    const correctResponses = results.responses.filter(r => r.isCorrect && !r.isDontKnow);
    const incorrectResponses = results.responses.filter(r => !r.isCorrect && !r.isDontKnow);
    const dontKnowResponses = results.responses.filter(r => r.isDontKnow);
    
    const avgCorrectRT = correctResponses.length > 0 ? 
        correctResponses.reduce((sum, r) => sum + r.reactionTime, 0) / correctResponses.length : 0;
    
    const avgIncorrectRT = incorrectResponses.length > 0 ? 
        incorrectResponses.reduce((sum, r) => sum + r.reactionTime, 0) / incorrectResponses.length : 0;
    
    const avgDontKnowRT = dontKnowResponses.length > 0 ? 
        dontKnowResponses.reduce((sum, r) => sum + r.reactionTime, 0) / dontKnowResponses.length : 0;
    
    // 結果の追加表示
    const detailedHTML = `
        <div style="margin-top: 30px;">
            <h3>反応パターン分析</h3>
            <p>正解の場合の平均反応時間: ${avgCorrectRT.toFixed(2)} 秒</p>
            <p>不正解の場合の平均反応時間: ${avgIncorrectRT.toFixed(2)} 秒</p>
            <p>「わからない」選択時の平均反応時間: ${avgDontKnowRT.toFixed(2)} 秒</p>
            
            <div style="margin-top: 15px;">
                <h4>反応パターンの解釈</h4>
                <p>通常、知っている単語（正解）には素早く反応し、わからない単語には時間がかかるか「わからない」を選択する傾向があります。</p>
                <p>不正解の反応時間が短い場合は推測（guessing）が働いている可能性があり、「わからない」の使用は推測を減少させることができます。</p>
            </div>
        </div>`;
    
    levelBreakdownElem.innerHTML += detailedHTML;
}

// 能力値チャート表示関数
function createAbilityChart() {
    // レベル別の正答率を計算
    const levelData = [];
    for (const level in results.levelScores) {
        const score = results.levelScores[level];
        const percentage = (score.correct / score.total * 100).toFixed(1);
        levelData.push({
            level: parseInt(level),
            correctRate: score.correct / score.total,
            percentage: percentage
        });
    }
    
    // レベル順にソート
    levelData.sort((a, b) => a.level - b.level);
    
    // データがなければ空のメッセージを返す
    if (levelData.length === 0) {
        return '<div class="chart-container"><p>データがありません</p></div>';
    }
    
    // SVGでチャートを作成
    const width = 500;
    const height = 300;
    const marginTop = 20;
    const marginRight = 30;
    const marginBottom = 40;
    const marginLeft = 50;
    
    // SVG要素を作成
    let chartHTML = `<div class="chart-container">
        <h3>レベル別正答率</h3>
        <svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" style="background-color: #f9f9f9; border-radius: 5px;">`;
    
    // グラフの実効領域
    const graphWidth = width - marginLeft - marginRight;
    const graphHeight = height - marginTop - marginBottom;
    
    // X軸とY軸のスケールを作成
    const xScale = graphWidth / (levelData.length > 1 ? levelData.length - 1 : 1);
    
    // X軸を描画
    chartHTML += `<line x1="${marginLeft}" y1="${height - marginBottom}" x2="${width - marginRight}" y2="${height - marginBottom}" stroke="black" stroke-width="1.5" />`;
    
    // Y軸を描画
    chartHTML += `<line x1="${marginLeft}" y1="${marginTop}" x2="${marginLeft}" y2="${height - marginBottom}" stroke="black" stroke-width="1.5" />`;
    
    // X軸のラベル
    levelData.forEach((data, i) => {
        const x = marginLeft + i * xScale;
        chartHTML += `<text x="${x}" y="${height - marginBottom + 20}" text-anchor="middle" font-size="12">L${data.level}</text>`;
    });
    
    // Y軸のラベル（正答率）
    for (let i = 0; i <= 10; i++) {
        const y = height - marginBottom - (i / 10) * graphHeight;
        const value = i * 10;
        chartHTML += `<text x="${marginLeft - 10}" y="${y}" text-anchor="end" alignment-baseline="middle" font-size="11">${value}%</text>`;
        // 水平グリッド線
        chartHTML += `<line x1="${marginLeft}" y1="${y}" x2="${width - marginRight}" y2="${y}" stroke="#ddd" stroke-width="1" stroke-dasharray="3,3" />`;
    }
    
    // 折れ線グラフのパスを作成
    if (levelData.length > 0) {
        chartHTML += `<polyline points="`;
        levelData.forEach((data, i) => {
            const x = marginLeft + i * xScale;
            const y = height - marginBottom - (data.correctRate * graphHeight);
            chartHTML += `${x},${y} `;
        });
        chartHTML += `" fill="none" stroke="#4a90e2" stroke-width="2.5" />`;
        
        // データポイントを描画
        levelData.forEach((data, i) => {
            const x = marginLeft + i * xScale;
            const y = height - marginBottom - (data.correctRate * graphHeight);
            // 点の下に白い円を描いて視認性を高める
            chartHTML += `<circle cx="${x}" cy="${y}" r="5" fill="white" stroke="#4a90e2" stroke-width="1.5" />`;
            chartHTML += `<circle cx="${x}" cy="${y}" r="3.5" fill="#4a90e2" />`;
            // パーセント値を表示
            chartHTML += `<text x="${x}" y="${y - 12}" text-anchor="middle" font-size="11" font-weight="bold">${data.percentage}%</text>`;
        });
    }
    
    chartHTML += `</svg></div>`;
    return chartHTML;
}

// 反応時間チャート表示関数
function createReactionTimeChart() {
    if (results.reactionTimes.length === 0) {
        return '<div class="chart-container"><p>反応時間データがありません</p></div>';
    }
    
    // レベル別の平均反応時間を計算
    const levelRTs = {};
    results.responses.forEach(response => {
        const level = testData[response.questionIndex].level;
        if (!levelRTs[level]) {
            levelRTs[level] = {
                total: 0,
                count: 0
            };
        }
        levelRTs[level].total += response.reactionTime;
        levelRTs[level].count++;
    });
    
    // 平均値を計算
    const rtData = [];
    for (const level in levelRTs) {
        const data = levelRTs[level];
        rtData.push({
            level: parseInt(level),
            averageRT: data.total / data.count
        });
    }
    
    // レベル順にソート
    rtData.sort((a, b) => a.level - b.level);
    
    // SVGでチャートを作成
    const width = 500;
    const height = 300;
    const marginTop = 20;
    const marginRight = 30;
    const marginBottom = 40;
    const marginLeft = 50;
    
    // グラフの実効領域
    const graphWidth = width - marginLeft - marginRight;
    const graphHeight = height - marginTop - marginBottom;
    
    // 最大RT値を取得（スケール用）
    const maxRT = rtData.length > 0 ? Math.max(...rtData.map(d => d.averageRT)) * 1.1 : 5; // 10%マージン
    
    // SVG要素を作成
    let chartHTML = `<div class="chart-container">
        <h3>レベル別平均反応時間</h3>
        <svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" style="background-color: #f9f9f9; border-radius: 5px;">`;
    
    // X軸とY軸のスケールを作成
    const xScale = graphWidth / (rtData.length > 1 ? rtData.length - 1 : 1);
    const yScale = graphHeight / maxRT;
    
    // X軸を描画
    chartHTML += `<line x1="${marginLeft}" y1="${height - marginBottom}" x2="${width - marginRight}" y2="${height - marginBottom}" stroke="black" stroke-width="1.5" />`;
    
    // Y軸を描画
    chartHTML += `<line x1="${marginLeft}" y1="${marginTop}" x2="${marginLeft}" y2="${height - marginBottom}" stroke="black" stroke-width="1.5" />`;
    
    // X軸のラベル
    rtData.forEach((data, i) => {
        const x = marginLeft + i * xScale;
        chartHTML += `<text x="${x}" y="${height - marginBottom + 20}" text-anchor="middle" font-size="12">L${data.level}</text>`;
    });
    
    // Y軸のラベル（秒数）
    const yTickCount = 5;
    for (let i = 0; i <= yTickCount; i++) {
        const value = (maxRT / yTickCount) * i;
        const y = height - marginBottom - value * yScale;
        chartHTML += `<text x="${marginLeft - 10}" y="${y}" text-anchor="end" alignment-baseline="middle" font-size="11">${value.toFixed(1)}秒</text>`;
        // 水平グリッド線
        chartHTML += `<line x1="${marginLeft}" y1="${y}" x2="${width - marginRight}" y2="${y}" stroke="#ddd" stroke-width="1" stroke-dasharray="3,3" />`;
    }
    
    // データがある場合のみ棒グラフを描画
    if (rtData.length > 0) {
        // 棒グラフを描画
        const barWidth = Math.min(xScale * 0.6, 60); // 最大幅を制限
        rtData.forEach((data, i) => {
            const x = marginLeft + i * xScale - barWidth/2;
            const barHeight = data.averageRT * yScale;
            const y = height - marginBottom - barHeight;
            
            // グラデーションを使った棒グラフ
            chartHTML += `<rect x="${x}" y="${y}" width="${barWidth}" height="${barHeight}" rx="2" ry="2" fill="#f39c12" opacity="0.85" />`;
            // 値のラベル
            chartHTML += `<text x="${marginLeft + i * xScale}" y="${y - 5}" text-anchor="middle" font-size="11" font-weight="bold">${data.averageRT.toFixed(2)}秒</text>`;
        });
    }
    
    chartHTML += `</svg></div>`;
    return chartHTML;
}

// 反応時間チャート表示関数
function createReactionTimeChart() {
    if (results.reactionTimes.length === 0) return "";
    
    // レベル別の平均反応時間を計算
    const levelRTs = {};
    results.responses.forEach(response => {
        const level = testData[response.questionIndex].level;
        if (!levelRTs[level]) {
            levelRTs[level] = {
                total: 0,
                count: 0
            };
        }
        levelRTs[level].total += response.reactionTime;
        levelRTs[level].count++;
    });
    
    // 平均値を計算
    const rtData = [];
    for (const level in levelRTs) {
        const data = levelRTs[level];
        rtData.push({
            level: parseInt(level),
            averageRT: data.total / data.count
        });
    }
    
    // レベル順にソート
    rtData.sort((a, b) => a.level - b.level);
    
    // SVGでチャートを作成
    const width = 500;
    const height = 300;
    const marginTop = 20;
    const marginRight = 30;
    const marginBottom = 40;
    const marginLeft = 50;
    
    // 最大RT値を取得（スケール用）- データが空の場合のフォールバック
    const maxRT = rtData.length > 0 ? Math.max(...rtData.map(d => d.averageRT)) * 1.1 : 10; // 10%マージン
    
    // SVG要素を作成
    let chartHTML = `<div class="chart-container">
        <h3>レベル別平均反応時間</h3>
        <svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">`;
    
    // X軸とY軸のスケールを作成
    const xScale = (width - marginLeft - marginRight) / (rtData.length - 1 || 1);
    const yScale = (height - marginTop - marginBottom) / maxRT;
    
    // X軸を描画
    chartHTML += `<line x1="${marginLeft}" y1="${height - marginBottom}" x2="${width - marginRight}" y2="${height - marginBottom}" stroke="black" />`;
    
    // Y軸を描画
    chartHTML += `<line x1="${marginLeft}" y1="${marginTop}" x2="${marginLeft}" y2="${height - marginBottom}" stroke="black" />`;
    
    // X軸のラベル
    rtData.forEach((data, i) => {
        const x = marginLeft + i * xScale;
        chartHTML += `<text x="${x}" y="${height - marginBottom + 20}" text-anchor="middle">L${data.level}</text>`;
    });
    
    // Y軸のラベル（秒数）
    for (let i = 0; i <= 5; i++) {
        const value = (maxRT / 5) * i;
        const y = height - marginBottom - value * yScale;
        chartHTML += `<text x="${marginLeft - 10}" y="${y}" text-anchor="end" alignment-baseline="middle">${value.toFixed(1)}s</text>`;
        chartHTML += `<line x1="${marginLeft - 5}" y1="${y}" x2="${marginLeft}" y2="${y}" stroke="black" />`;
    }
    
    // データがある場合のみ棒グラフを描画
    if (rtData.length > 0) {
        // 棒グラフを描画
        const barWidth = xScale * 0.6;
        rtData.forEach((data, i) => {
            const x = marginLeft + i * xScale - barWidth/2;
            const barHeight = data.averageRT * yScale;
            const y = height - marginBottom - barHeight;
            
            chartHTML += `<rect x="${x}" y="${y}" width="${barWidth}" height="${barHeight}" fill="#f39c12" />`;
            chartHTML += `<text x="${marginLeft + i * xScale}" y="${y - 5}" text-anchor="middle" font-size="12">${data.averageRT.toFixed(2)}s</text>`;
        });
    }
    
    chartHTML += `</svg></div>`;
    return chartHTML;
}

// 3パラメタロジスティックモデル(3PLM)を使った語彙サイズ推定関数
function estimateVocabularySize() {
    // 能力値（theta）の推定
    const estimatedTheta = estimateAbility();
    
    // 日本人英語学習者の語彙サイズ推定モデル（仮の値）
    // VST-NJ8の各レベルが対応する語彙レベルを仮定
    // レベル1: 1000語レベル、レベル2: 2000語レベル...レベル8: 8000語レベル
    const baseVocabSize = 1000;
    
    // 能力値からおおよその語彙サイズを推定
    // theta = 0 を4000語レベルとし、1の変化につき約2000語の変化と仮定
    const estimatedSize = baseVocabSize * (4 + estimatedTheta * 2);
    
    // 負の値にならないよう、また最大8000語を超えないよう調整
    return Math.max(Math.min(Math.round(estimatedSize), 8000), 0);
}

// 項目反応理論(IRT)に基づく能力値推定関数
function estimateAbility() {
    // 最尤推定法（Maximum Likelihood Estimation）によるtheta推定
    const MAX_ITERATIONS = 20;
    const CONVERGENCE_CRITERION = 0.001;
    
    // 初期値を設定
    let theta = 0;
    
    // 反応データがない場合は初期値を返す
    if (results.responses.length === 0) {
        return theta;
    }
    
    try {
        // ニュートン・ラフソン法による反復
        for (let iter = 0; iter < MAX_ITERATIONS; iter++) {
            // 勾配と曲率を計算
            let gradient = 0;
            let curvature = 0;
            
            // 各問題への反応データを使って計算
            for (let i = 0; i < results.responses.length; i++) {
                const response = results.responses[i];
                const item = testData[response.questionIndex];
                
                // 3PLMにおけるパラメータ
                const a = item.discrimination;
                const b = item.difficulty;
                const c = item.guessing;
                
                // 正解確率を計算
                const z = a * (theta - b);
                const p = c + (1 - c) / (1 + Math.exp(-z));
                const pPrime = (1 - c) * a * Math.exp(-z) / Math.pow(1 + Math.exp(-z), 2);
                const pDoublePrime = (1 - c) * a * a * Math.exp(-z) * (Math.exp(-z) - 1) / Math.pow(1 + Math.exp(-z), 3);
                
                // 「わからない」を選択した場合は0.2（ランダム解答の確率）とみなす
                const u = response.isDontKnow ? 0.2 : (response.isCorrect ? 1 : 0);
                
                // 勾配と曲率を更新
                gradient += (u - p) * pPrime / (p * (1 - p));
                curvature += Math.pow(pPrime, 2) / (p * (1 - p)) - (u - p) * pDoublePrime / (p * (1 - p));
            }
            
            // theta更新
            const deltaTheta = gradient / curvature;
            theta += deltaTheta;
            
            // 収束判定
            if (Math.abs(deltaTheta) < CONVERGENCE_CRITERION) {
                break;
            }
        }
        
        // 特殊なケース（全問正解/全問不正解）の処理
        if (results.correctCount === 0) {
            theta = -4.0; // 非常に低い能力値
        } else if (results.correctCount === testData.length) {
            theta = 4.0;  // 非常に高い能力値
        }
        
        // 異常値のチェック
        if (isNaN(theta) || !isFinite(theta)) {
            return 0; // デフォルト値
        }
        
        // 範囲の制限
        return Math.max(Math.min(theta, 4.0), -4.0);
    } catch (error) {
        console.error("能力値推定中にエラーが発生しました:", error);
        return 0; // エラーが発生した場合はデフォルト値
    }
}

// 配列をシャッフルする関数
function shuffleArray(array) {
    const newArray = [...array];
    for (let i = newArray.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [newArray[i], newArray[j]] = [newArray[j], newArray[i]];
    }
    return newArray;
}