using System.Diagnostics;
//using MyExcelMAUIApp;

namespace DemoParser.Parsing;
public class LabCalculatorVisitor : LabCalculatorBaseVisitor<double>
{
    private readonly IDictionary<string, double> tableIdentifier;

    public LabCalculatorVisitor(IDictionary<string, double> tableIdentifier)
    {
        ArgumentNullException.ThrowIfNull(tableIdentifier);
        this.tableIdentifier=tableIdentifier;
    }
    
    public override double VisitCompileUnit(LabCalculatorParser.CompileUnitContext context)
    {
        return Visit(context.expression());
    }

    public override double VisitNumberExpr(LabCalculatorParser.NumberExprContext context)
    {
        var result = double.Parse(context.GetText());
        Debug.WriteLine(result);

        return result;
    }

    //IdentifierExpr
    public override double VisitIdentifierExpr(LabCalculatorParser.IdentifierExprContext context)
    {
        var result = context.GetText();
        //видобути значення змінної з таблиці
        if (tableIdentifier.TryGetValue(result.ToString(), out double value))
        {
            return value;
        }
        else
        {
            return 0.0;
        }
    }

    public override double VisitParenthesizedExpr(LabCalculatorParser.ParenthesizedExprContext context)
    {
        return Visit(context.expression());
    }

    public override double VisitExponentialExpr(LabCalculatorParser.ExponentialExprContext context)
    {
        var left = WalkLeft(context);
        var right = WalkRight(context);

        Debug.WriteLine("{0} ^ {1}", left, right);
        return System.Math.Pow(left, right);
    }

    public override double VisitIncExpr(LabCalculatorParser.IncExprContext context)
    {
        var left = WalkLeft(context);

        if (context.operatorToken.Type == LabCalculatorLexer.INC)
        {
            Debug.WriteLine("inc({0})", left);
            return left + 1;
        }
        else //LabCalculatorLexer.DEC
        {
            Debug.WriteLine("dec({0})", left);
            return left - 1;
        }
    }

    public override double VisitMaxExpr(LabCalculatorParser.MaxExprContext context)
    {
        var left = WalkLeft(context);
        var right = WalkRight(context);

        if (context.operatorToken.Type == LabCalculatorLexer.MAX)
        {
            Debug.WriteLine("max( {0}, {1} )", left, right);
            return System.Math.Max(left, right);
        }
        else //LabCalculatorLexer.MIN
        {
            Debug.WriteLine("min( {0}, {1} )", left, right);
            return System.Math.Min(left, right);
        }
    }

    public override double VisitAdditiveExpr(LabCalculatorParser.AdditiveExprContext context)
    {
        var left = WalkLeft(context);
        var right = WalkRight(context);

        if (context.operatorToken.Type == LabCalculatorLexer.ADD)
        {
            Debug.WriteLine("{0} + {1}", left, right);
            return left + right;
        }
        else //LabCalculatorLexer.SUBTRACT
        {
            Debug.WriteLine("{0} - {1}", left, right);
            return left - right;
        }
    }

    public override double VisitMultiplicativeExpr(LabCalculatorParser.MultiplicativeExprContext context)
    {
        var left = WalkLeft(context);
        var right = WalkRight(context);

        if (context.operatorToken.Type == LabCalculatorLexer.MULTIPLY)
        {
            Debug.WriteLine("{0} * {1}", left, right);
            return left * right;
        }
        else //LabCalculatorLexer.DIVIDE
        {
            Debug.WriteLine("{0} / {1}", left, right);
            return left / right;
        }
    }

    private double WalkLeft(LabCalculatorParser.ExpressionContext context)
    {
        return Visit(context.GetRuleContext<LabCalculatorParser.ExpressionContext>(0));
    }

    private double WalkRight(LabCalculatorParser.ExpressionContext context)
    {
        return Visit(context.GetRuleContext<LabCalculatorParser.ExpressionContext>(1));
    }
}

