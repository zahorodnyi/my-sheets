using System.Globalization;
using System.Text;
using MySheets.Core.Common;

namespace MySheets.Core.Calculation;

public class FormulaEvaluator {
    private static readonly HashSet<string> Functions = new() { "SUM", "AVERAGE", "MAX", "MIN", "MEDIAN" };

    public object Evaluate(string expression, Func<string, object> getVariableValue) {
        if (string.IsNullOrEmpty(expression)) return string.Empty;
        
        if (expression.StartsWith('=')) {
            expression = expression.Substring(1);
        }

        try {
            var tokens = Tokenize(expression);
            var rpn = ShuntingYard(tokens);
            return EvaluateRPN(rpn, getVariableValue);
        } 
        catch {
            return "#ERROR!";
        }
    }

    private List<string> Tokenize(string expression) {
        var tokens = new List<string>();
        var buffer = new StringBuilder();

        for (int i = 0; i < expression.Length; i++) {
            char c = expression[i];

            if (char.IsWhiteSpace(c)) {
                if (buffer.Length > 0) AddToken(tokens, buffer);
                continue;
            }

            if (c == ':') {
                if (buffer.Length > 0) {
                    string startRef = buffer.ToString();
                    buffer.Clear();
                    
                    var endRefBuilder = new StringBuilder();
                    int j = i + 1;
                    while (j < expression.Length && char.IsLetterOrDigit(expression[j])) {
                        endRefBuilder.Append(expression[j]);
                        j++;
                    }
                    
                    if (endRefBuilder.Length > 0) {
                        string endRef = endRefBuilder.ToString();
                        ExpandRange(tokens, startRef, endRef);
                        i = j - 1; 
                        continue;
                    } 
                    else {
                        tokens.Add(startRef.ToUpper());
                        tokens.Add(":");
                        continue;
                    }
                }
            }

            if (c == '+') {
                if (buffer.Length == 0) {
                    if (tokens.Count == 0) continue; 
                    var lastToken = tokens[tokens.Count - 1];
                    if (lastToken == "(" || lastToken == "," || IsOperator(lastToken)) continue; 
                }
            }

            if (c == '-') {
                bool isNegativeNumber = false;
                if (i + 1 < expression.Length && char.IsDigit(expression[i + 1])) {
                    if (buffer.Length == 0) {
                        if (tokens.Count == 0) {
                            isNegativeNumber = true;
                        } 
                        else {
                            var lastToken = tokens[tokens.Count - 1];
                            if (lastToken == "(" || lastToken == "," || IsOperator(lastToken)) {
                                isNegativeNumber = true;
                            }
                        }
                    }
                }

                if (isNegativeNumber) {
                    buffer.Append(c);
                    continue; 
                }
            }

            if (char.IsDigit(c) || c == '.') {
                buffer.Append(c);
            } 
            else if (char.IsLetter(c)) {
                if (buffer.Length > 0 && IsNumeric(buffer.ToString())) {
                    AddToken(tokens, buffer);
                }
                
                buffer.Append(c);
                while (i + 1 < expression.Length && (char.IsLetterOrDigit(expression[i + 1]))) {
                    buffer.Append(expression[++i]);
                }
                
                if (i + 1 < expression.Length && expression[i+1] == ':') {
                    continue; 
                }

                AddToken(tokens, buffer);
            } 
            else {
                if (buffer.Length > 0) {
                    AddToken(tokens, buffer);
                }
                tokens.Add(c.ToString());
            }
        }

        if (buffer.Length > 0) {
            AddToken(tokens, buffer);
        }

        return tokens;
    }

    private void ExpandRange(List<string> tokens, string start, string end) {
        try {
            var range = $"{start}:{end}";
            var cells = CellReferenceUtility.ParseRange(range).ToList();
            
            for (int k = 0; k < cells.Count; k++) {
                var (r, c) = cells[k];
                string colName = CellReferenceUtility.GetColumnName(c);
                tokens.Add($"{colName}{r + 1}");
                
                if (k < cells.Count - 1) {
                    tokens.Add(","); 
                }
            }
        } 
        catch {
            tokens.Add(start.ToUpper());
            tokens.Add(":");
            tokens.Add(end.ToUpper());
        }
    }

    private void AddToken(List<string> tokens, StringBuilder buffer) {
        tokens.Add(buffer.ToString().ToUpper());
        buffer.Clear();
    }

    private bool IsNumeric(string s) => double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out _);

    private Queue<string> ShuntingYard(List<string> tokens) {
        var output = new Queue<string>();
        var operators = new Stack<string>();
        var argCounts = new Stack<int>();

        for (int i = 0; i < tokens.Count; i++) {
            var token = tokens[i];

            if (double.TryParse(token, NumberStyles.Any, CultureInfo.InvariantCulture, out _)) {
                output.Enqueue(token);
            } 
            else if (Functions.Contains(token)) {
                operators.Push(token);
                argCounts.Push(1);
            } 
            else if (token == ",") {
                while (operators.Count > 0 && operators.Peek() != "(") {
                    output.Enqueue(operators.Pop());
                }
                if (argCounts.Count > 0) {
                    argCounts.Push(argCounts.Pop() + 1);
                }
            } 
            else if (token == "(") {
                operators.Push(token);
            } 
            else if (token == ")") {
                while (operators.Count > 0 && operators.Peek() != "(") {
                    output.Enqueue(operators.Pop());
                }
                
                if (operators.Count == 0) throw new InvalidOperationException("Mismatched parentheses");
                operators.Pop();

                if (operators.Count > 0 && Functions.Contains(operators.Peek())) {
                    var func = operators.Pop();
                    var args = argCounts.Pop();
                    
                    if (tokens[i - 1] == "(") args = 0;
                    
                    output.Enqueue($"{func}:{args}");
                }
            } 
            else if (IsIdentifier(token)) {
                output.Enqueue(token);
            } 
            else if (IsOperator(token)) {
                while (operators.Count > 0 && Precedence(operators.Peek()) >= Precedence(token)) {
                    output.Enqueue(operators.Pop());
                }
                operators.Push(token);
            }
        }

        while (operators.Count > 0) {
            if (operators.Peek() == "(") throw new InvalidOperationException("Mismatched parentheses");
            output.Enqueue(operators.Pop());
        }

        return output;
    }

    private object EvaluateRPN(Queue<string> tokens, Func<string, object> getVariableValue) {
        var stack = new Stack<object>();

        while (tokens.Count > 0) {
            var token = tokens.Dequeue();

            if (double.TryParse(token, NumberStyles.Any, CultureInfo.InvariantCulture, out double value)) {
                stack.Push(value);
            } 
            else if (token.Contains(':')) {
                var parts = token.Split(':');
                var funcName = parts[0];
                var argCount = int.Parse(parts[1]);
                var args = new List<double>();

                for (int i = 0; i < argCount; i++) {
                    if (stack.Count == 0) throw new InvalidOperationException();
                    var obj = stack.Pop();
                    if (IsCycleError(obj)) { stack.Push("#CYCLE!"); goto NextToken; }
                    args.Add(Convert.ToDouble(obj, CultureInfo.InvariantCulture));
                }
                
                if (args.Count == 0 && funcName != "COUNT") throw new InvalidOperationException();

                switch (funcName) {
                    case "SUM": stack.Push(args.Sum()); break;
                    case "AVERAGE": stack.Push(args.Average()); break;
                    case "MAX": stack.Push(args.Max()); break;
                    case "MIN": stack.Push(args.Min()); break;
                    case "MEDIAN": 
                        args.Sort();
                        int count = args.Count;
                        if (count % 2 == 0) {
                            stack.Push((args[count / 2 - 1] + args[count / 2]) / 2.0);
                        }
                        else {
                            stack.Push(args[count / 2]);
                        }

                        break;
                }
                NextToken:;
            } 
            else if (IsIdentifier(token)) {
                var val = getVariableValue(token);
                stack.Push(val);
            } 
            else {
                if (stack.Count < 2) throw new InvalidOperationException("Invalid expression");
                
                var rightObj = stack.Pop();
                var leftObj = stack.Pop();

                if (IsCycleError(leftObj)) { stack.Push("#CYCLE!"); continue; }
                if (IsCycleError(rightObj)) { stack.Push("#CYCLE!"); continue; }

                double right = Convert.ToDouble(rightObj, CultureInfo.InvariantCulture);
                double left = Convert.ToDouble(leftObj, CultureInfo.InvariantCulture);

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

    private bool IsCycleError(object obj) {
        return obj is string s && s == "#CYCLE!";
    }

    private bool IsIdentifier(string token) {
        return char.IsLetter(token[0]) && !Functions.Contains(token);
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