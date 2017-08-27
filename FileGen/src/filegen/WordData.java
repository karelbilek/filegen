package filegen;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class WordData {
    public Map<String, String> strings;
    public Map<String, List<String>> lists;

    public Map<String,Object> all() {
        Map<String, Object> r = new HashMap<>();
        r.putAll(strings);
        r.putAll(lists);
        return r;
    }
}
