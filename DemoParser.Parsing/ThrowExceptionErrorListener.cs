using Antlr4.Runtime;

namespace DemoParser.Parsing;

public class ThrowExceptionErrorListener : BaseErrorListener, IAntlrErrorListener<int>
{
    //BaseErrorListener implementation
    public void SyntaxError(TextWriter output, IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
    {
        throw new ArgumentException("Неправильний вираз: {0}", msg, e);
    }

    //IAntlrErrorListener<int> implementation
    public void SyntaxError(IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
    {
        throw new ArgumentException("Неправильний вираз: {0}", msg, e);
    }

    
}
