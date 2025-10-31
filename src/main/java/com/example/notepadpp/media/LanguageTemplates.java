package com.example.notepadpp.media;
import org.fife.ui.rsyntaxtextarea.SyntaxConstants;
public class LanguageTemplates {
    
    public static String getTemplate(String language) {
        switch (language) {
            case "Java":
                return "import java.util.Scanner;\n\npublic class MyProgram {\n    public static void main(String[] args) {\n        Scanner scanner = new Scanner(System.in);\n        System.out.print(\"Enter your name: \");\n        String name = scanner.nextLine();\n        System.out.println(\"Hello, \" + name + \"!\");\n        System.out.print(\"Enter another name: \");\n        String name2 = scanner.nextLine();\n        System.out.println(\"Hello again, \" + name2 + \"!\");\n    }\n}";
            
            case "Python":
                return "name = input(\"Enter your name: \")\nprint(f\"Hello, {name}!\")\nname2 = input(\"Enter another name: \")\nprint(f\"Hello again, {name2}!\")";
            
            case "C++":
                return "#include <iostream>\n#include <string>\n\nint main() {\n    std::string name;\n    std::cout << \"Enter your name: \";\n    std::getline(std::cin, name);\n    std::cout << \"Hello, \" << name << \"!\" << std::endl;\n    \n    std::string name2;\n    std::cout << \"Enter another name: \";\n    std::getline(std::cin, name2);\n    std::cout << \"Hello again, \" << name2 << \"!\" << std::endl;\n    \n    return 0;\n}";
            
            case "C":
                return "#include <stdio.h>\n\nint main() {\n    char name[100];\n    printf(\"Enter your name: \");\n    fgets(name, sizeof(name), stdin);\n    printf(\"Hello, %s\", name);\n    \n    char name2[100];\n    printf(\"Enter another name: \");\n    fgets(name2, sizeof(name2), stdin);\n    printf(\"Hello again, %s\", name2);\n    \n    return 0;\n}";
            
            default:
                return "// Unknown language template";
        }
    }
    
    public static String getSyntaxStyle(String language) {
        switch (language) {
            case "Java": return SyntaxConstants.SYNTAX_STYLE_JAVA;
            case "Python": return SyntaxConstants.SYNTAX_STYLE_PYTHON;
            case "C++": return SyntaxConstants.SYNTAX_STYLE_CPLUSPLUS;
            case "C": return SyntaxConstants.SYNTAX_STYLE_C;
            default: return SyntaxConstants.SYNTAX_STYLE_NONE;
        }
    }
}
