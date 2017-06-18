import java.util.Scanner;

public class PhraseRepeaterDriver{
      public static void main (String[] args) {
          Scanner kb = new Scanner(System.in);

          System.out.print("Enter a message: ");
          String msg = kb.nextLine();
          System.out.print("Number of times: ");
          int n = kb.nextInt();

          PhraseRepeater pr = new PhraseRepeater();
          pr.setValue(msg, n);
          System.out.println(pr.getRepaeatedPhrase());
      }
}
