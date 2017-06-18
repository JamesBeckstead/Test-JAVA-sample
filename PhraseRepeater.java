public class PhraseRepeater{

  String phrase;
  int repeats;

      public void setValue( String p, int r) {
          phrase = p;
          repeats = r;
      }

      public String getRepaeatedPhrase() {
          String results = "";
          for (int i=0; i<repeats; i++)
              results += phrase;
          return results;
      }
}
