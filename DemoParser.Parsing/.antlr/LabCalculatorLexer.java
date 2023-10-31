// Generated from d:/Áóðñ³òºò/2_course/ÎÎÏ/MyExcelMAUIApp/DemoParser.Parsing/LabCalculator.g4 by ANTLR 4.13.1
import org.antlr.v4.runtime.Lexer;
import org.antlr.v4.runtime.CharStream;
import org.antlr.v4.runtime.Token;
import org.antlr.v4.runtime.TokenStream;
import org.antlr.v4.runtime.*;
import org.antlr.v4.runtime.atn.*;
import org.antlr.v4.runtime.dfa.DFA;
import org.antlr.v4.runtime.misc.*;

@SuppressWarnings({"all", "warnings", "unchecked", "unused", "cast", "CheckReturnValue", "this-escape"})
public class LabCalculatorLexer extends Lexer {
	static { RuntimeMetaData.checkVersion("4.13.1", RuntimeMetaData.VERSION); }

	protected static final DFA[] _decisionToDFA;
	protected static final PredictionContextCache _sharedContextCache =
		new PredictionContextCache();
	public static final int
		T__0=1, NUMBER=2, IDENTIFIER=3, INT=4, EXPONENT=5, MULTIPLY=6, DIVIDE=7, 
		SUBTRACT=8, ADD=9, LPAREN=10, RPAREN=11, INC=12, DEC=13, MAX=14, MIN=15, 
		WS=16;
	public static String[] channelNames = {
		"DEFAULT_TOKEN_CHANNEL", "HIDDEN"
	};

	public static String[] modeNames = {
		"DEFAULT_MODE"
	};

	private static String[] makeRuleNames() {
		return new String[] {
			"T__0", "NUMBER", "IDENTIFIER", "INT", "EXPONENT", "MULTIPLY", "DIVIDE", 
			"SUBTRACT", "ADD", "LPAREN", "RPAREN", "INC", "DEC", "MAX", "MIN", "WS"
		};
	}
	public static final String[] ruleNames = makeRuleNames();

	private static String[] makeLiteralNames() {
		return new String[] {
			null, "','", null, null, null, "'^'", "'*'", "'/'", "'-'", "'+'", "'('", 
			"')'", "'inc'", "'dec'", "'max'", "'min'"
		};
	}
	private static final String[] _LITERAL_NAMES = makeLiteralNames();
	private static String[] makeSymbolicNames() {
		return new String[] {
			null, null, "NUMBER", "IDENTIFIER", "INT", "EXPONENT", "MULTIPLY", "DIVIDE", 
			"SUBTRACT", "ADD", "LPAREN", "RPAREN", "INC", "DEC", "MAX", "MIN", "WS"
		};
	}
	private static final String[] _SYMBOLIC_NAMES = makeSymbolicNames();
	public static final Vocabulary VOCABULARY = new VocabularyImpl(_LITERAL_NAMES, _SYMBOLIC_NAMES);

	/**
	 * @deprecated Use {@link #VOCABULARY} instead.
	 */
	@Deprecated
	public static final String[] tokenNames;
	static {
		tokenNames = new String[_SYMBOLIC_NAMES.length];
		for (int i = 0; i < tokenNames.length; i++) {
			tokenNames[i] = VOCABULARY.getLiteralName(i);
			if (tokenNames[i] == null) {
				tokenNames[i] = VOCABULARY.getSymbolicName(i);
			}

			if (tokenNames[i] == null) {
				tokenNames[i] = "<INVALID>";
			}
		}
	}

	@Override
	@Deprecated
	public String[] getTokenNames() {
		return tokenNames;
	}

	@Override

	public Vocabulary getVocabulary() {
		return VOCABULARY;
	}


	public LabCalculatorLexer(CharStream input) {
		super(input);
		_interp = new LexerATNSimulator(this,_ATN,_decisionToDFA,_sharedContextCache);
	}

	@Override
	public String getGrammarFileName() { return "LabCalculator.g4"; }

	@Override
	public String[] getRuleNames() { return ruleNames; }

	@Override
	public String getSerializedATN() { return _serializedATN; }

	@Override
	public String[] getChannelNames() { return channelNames; }

	@Override
	public String[] getModeNames() { return modeNames; }

	@Override
	public ATN getATN() { return _ATN; }

	public static final String _serializedATN =
		"\u0004\u0000\u0010[\u0006\uffff\uffff\u0002\u0000\u0007\u0000\u0002\u0001"+
		"\u0007\u0001\u0002\u0002\u0007\u0002\u0002\u0003\u0007\u0003\u0002\u0004"+
		"\u0007\u0004\u0002\u0005\u0007\u0005\u0002\u0006\u0007\u0006\u0002\u0007"+
		"\u0007\u0007\u0002\b\u0007\b\u0002\t\u0007\t\u0002\n\u0007\n\u0002\u000b"+
		"\u0007\u000b\u0002\f\u0007\f\u0002\r\u0007\r\u0002\u000e\u0007\u000e\u0002"+
		"\u000f\u0007\u000f\u0001\u0000\u0001\u0000\u0001\u0001\u0001\u0001\u0001"+
		"\u0001\u0003\u0001\'\b\u0001\u0001\u0002\u0004\u0002*\b\u0002\u000b\u0002"+
		"\f\u0002+\u0001\u0002\u0001\u0002\u0005\u00020\b\u0002\n\u0002\f\u0002"+
		"3\t\u0002\u0001\u0003\u0004\u00036\b\u0003\u000b\u0003\f\u00037\u0001"+
		"\u0004\u0001\u0004\u0001\u0005\u0001\u0005\u0001\u0006\u0001\u0006\u0001"+
		"\u0007\u0001\u0007\u0001\b\u0001\b\u0001\t\u0001\t\u0001\n\u0001\n\u0001"+
		"\u000b\u0001\u000b\u0001\u000b\u0001\u000b\u0001\f\u0001\f\u0001\f\u0001"+
		"\f\u0001\r\u0001\r\u0001\r\u0001\r\u0001\u000e\u0001\u000e\u0001\u000e"+
		"\u0001\u000e\u0001\u000f\u0001\u000f\u0001\u000f\u0001\u000f\u0000\u0000"+
		"\u0010\u0001\u0001\u0003\u0002\u0005\u0003\u0007\u0004\t\u0005\u000b\u0006"+
		"\r\u0007\u000f\b\u0011\t\u0013\n\u0015\u000b\u0017\f\u0019\r\u001b\u000e"+
		"\u001d\u000f\u001f\u0010\u0001\u0000\u0004\u0002\u0000AZaz\u0001\u0000"+
		"19\u0001\u000009\u0003\u0000\t\n\r\r  ^\u0000\u0001\u0001\u0000\u0000"+
		"\u0000\u0000\u0003\u0001\u0000\u0000\u0000\u0000\u0005\u0001\u0000\u0000"+
		"\u0000\u0000\u0007\u0001\u0000\u0000\u0000\u0000\t\u0001\u0000\u0000\u0000"+
		"\u0000\u000b\u0001\u0000\u0000\u0000\u0000\r\u0001\u0000\u0000\u0000\u0000"+
		"\u000f\u0001\u0000\u0000\u0000\u0000\u0011\u0001\u0000\u0000\u0000\u0000"+
		"\u0013\u0001\u0000\u0000\u0000\u0000\u0015\u0001\u0000\u0000\u0000\u0000"+
		"\u0017\u0001\u0000\u0000\u0000\u0000\u0019\u0001\u0000\u0000\u0000\u0000"+
		"\u001b\u0001\u0000\u0000\u0000\u0000\u001d\u0001\u0000\u0000\u0000\u0000"+
		"\u001f\u0001\u0000\u0000\u0000\u0001!\u0001\u0000\u0000\u0000\u0003#\u0001"+
		"\u0000\u0000\u0000\u0005)\u0001\u0000\u0000\u0000\u00075\u0001\u0000\u0000"+
		"\u0000\t9\u0001\u0000\u0000\u0000\u000b;\u0001\u0000\u0000\u0000\r=\u0001"+
		"\u0000\u0000\u0000\u000f?\u0001\u0000\u0000\u0000\u0011A\u0001\u0000\u0000"+
		"\u0000\u0013C\u0001\u0000\u0000\u0000\u0015E\u0001\u0000\u0000\u0000\u0017"+
		"G\u0001\u0000\u0000\u0000\u0019K\u0001\u0000\u0000\u0000\u001bO\u0001"+
		"\u0000\u0000\u0000\u001dS\u0001\u0000\u0000\u0000\u001fW\u0001\u0000\u0000"+
		"\u0000!\"\u0005,\u0000\u0000\"\u0002\u0001\u0000\u0000\u0000#&\u0003\u0007"+
		"\u0003\u0000$%\u0005.\u0000\u0000%\'\u0003\u0007\u0003\u0000&$\u0001\u0000"+
		"\u0000\u0000&\'\u0001\u0000\u0000\u0000\'\u0004\u0001\u0000\u0000\u0000"+
		"(*\u0007\u0000\u0000\u0000)(\u0001\u0000\u0000\u0000*+\u0001\u0000\u0000"+
		"\u0000+)\u0001\u0000\u0000\u0000+,\u0001\u0000\u0000\u0000,-\u0001\u0000"+
		"\u0000\u0000-1\u0007\u0001\u0000\u0000.0\u0007\u0002\u0000\u0000/.\u0001"+
		"\u0000\u0000\u000003\u0001\u0000\u0000\u00001/\u0001\u0000\u0000\u0000"+
		"12\u0001\u0000\u0000\u00002\u0006\u0001\u0000\u0000\u000031\u0001\u0000"+
		"\u0000\u000046\u000209\u000054\u0001\u0000\u0000\u000067\u0001\u0000\u0000"+
		"\u000075\u0001\u0000\u0000\u000078\u0001\u0000\u0000\u00008\b\u0001\u0000"+
		"\u0000\u00009:\u0005^\u0000\u0000:\n\u0001\u0000\u0000\u0000;<\u0005*"+
		"\u0000\u0000<\f\u0001\u0000\u0000\u0000=>\u0005/\u0000\u0000>\u000e\u0001"+
		"\u0000\u0000\u0000?@\u0005-\u0000\u0000@\u0010\u0001\u0000\u0000\u0000"+
		"AB\u0005+\u0000\u0000B\u0012\u0001\u0000\u0000\u0000CD\u0005(\u0000\u0000"+
		"D\u0014\u0001\u0000\u0000\u0000EF\u0005)\u0000\u0000F\u0016\u0001\u0000"+
		"\u0000\u0000GH\u0005i\u0000\u0000HI\u0005n\u0000\u0000IJ\u0005c\u0000"+
		"\u0000J\u0018\u0001\u0000\u0000\u0000KL\u0005d\u0000\u0000LM\u0005e\u0000"+
		"\u0000MN\u0005c\u0000\u0000N\u001a\u0001\u0000\u0000\u0000OP\u0005m\u0000"+
		"\u0000PQ\u0005a\u0000\u0000QR\u0005x\u0000\u0000R\u001c\u0001\u0000\u0000"+
		"\u0000ST\u0005m\u0000\u0000TU\u0005i\u0000\u0000UV\u0005n\u0000\u0000"+
		"V\u001e\u0001\u0000\u0000\u0000WX\u0007\u0003\u0000\u0000XY\u0001\u0000"+
		"\u0000\u0000YZ\u0006\u000f\u0000\u0000Z \u0001\u0000\u0000\u0000\u0005"+
		"\u0000&+17\u0001\u0000\u0001\u0000";
	public static final ATN _ATN =
		new ATNDeserializer().deserialize(_serializedATN.toCharArray());
	static {
		_decisionToDFA = new DFA[_ATN.getNumberOfDecisions()];
		for (int i = 0; i < _ATN.getNumberOfDecisions(); i++) {
			_decisionToDFA[i] = new DFA(_ATN.getDecisionState(i), i);
		}
	}
}