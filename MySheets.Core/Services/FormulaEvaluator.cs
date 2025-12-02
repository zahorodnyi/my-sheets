using System.Globalization;
using System.Text;

namespace MySheets.Core.Services;

public class FormulaEvaluator {
    public object Evaluate(string expression, Func<string, double> getVariableValue) {
        if (string.IsNullOrEmpty(expression)) return string.Empty;
        
        if (expression.StartsWith('=')) {
            expression = expression.Substring(1);
        }

        try {
            var tokens = Tokenize(expression);
            var rpn = ShuntingYard(tokens);
            return EvaluateRPN(rpn, getVariableValue);
        } catch {
            return "#ERROR!";
        }
    }

    private List<string> Tokenize(string expression) {
        var tokens = new List<string>();
        var buffer = new StringBuilder();

        for (int i = 0; i < expression.Length; i++) {
            char c = expression[i];

            if (char.IsWhiteSpace(c)) continue;

            if (char.IsDigit(c) || c == '.') {
                buffer.Append(c);
            } else if (char.IsLetter(c)) {
                buffer.Append(c);
                while (i + 1 < expression.Length && (char.IsLetterOrDigit(expression[i + 1]))) {
                    buffer.Append(expression[++i]);
                }
                tokens.Add(buffer.ToString());
                buffer.Clear();
            } else {
                if (buffer.Length > 0) {
                    tokens.Add(buffer.ToString());
                    buffer.Clear();
                }
                tokens.Add(c.ToString());
            }
        }

        if (buffer.Length > 0) {
            tokens.Add(buffer.ToString());
        }

        return tokens;
    }

    private Queue<string> ShuntingYard(List<string> tokens) {
        var output = new Queue<string>();
        var operators = new Stack<string>();

        foreach (var token in tokens) {
            if (double.TryParse(token, NumberStyles.Any, CultureInfo.InvariantCulture, out _) || IsIdentifier(token)) {
                output.Enqueue(token);
            } else if (token == "(") {
                operators.Push(token);
            } else if (token == ")") {
                while (operators.Count > 0 && operators.Peek() != "(") {
                    output.Enqueue(operators.Pop());
                }
                if (operators.Count > 0) operators.Pop();
            } else if (IsOperator(token)) {
                while (operators.Count > 0 && Precedence(operators.Peek()) >= Precedence(token)) {
                    output.Enqueue(operators.Pop());
                }
                operators.Push(token);
            }
        }

        while (operators.Count > 0) {
            output.Enqueue(operators.Pop());
        }

        return output;
    }

    private double EvaluateRPN(Queue<string> tokens, Func<string, double> getVariableValue) {
        var stack = new Stack<double>();

        while (tokens.Count > 0) {
            var token = tokens.Dequeue();

            if (double.TryParse(token, NumberStyles.Any, CultureInfo.InvariantCulture, out double value)) {
                stack.Push(value);
            } else if (IsIdentifier(token)) {
                stack.Push(getVariableValue(token));
            } else {
                if (stack.Count < 2) throw new InvalidOperationException("Invalid expression");
                
                var right = stack.Pop();
                var left = stack.Pop();

                switch (token) {
                    case "+": stack.Push(left + right); break;
                    case "-": stack.Push(left - right); break;
                    case "*": stack.Push(left * right); break;
                    case "/": 
                        if (right == 0) throw new DivideByZeroException();
                        stack.Push(left / right); 
                        break;
                    default: throw new InvalidOperationException("Unknown operator");
                }
            }
        }

        if (stack.Count != 1) throw new InvalidOperationException("Invalid expression");

        return stack.Pop();
    }

    private bool IsIdentifier(string token) {
        return char.IsLetter(token[0]);
    }

    private bool IsOperator(string token) {
        return token == "+" || token == "-" || token == "*" || token == "/";
    }

    private int Precedence(string op) {
        if (op == "*" || op == "/") return 2;
        if (op == "+" || op == "-") return 1;
        return 0;
    }
}