import com.paresh.diff.dto.DiffResponse;
import com.paresh.diff.renderer.ExcelRenderer;
import com.paresh.diff.util.DiffComputeEngine;
import org.junit.BeforeClass;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

public class ExcelRendererTest {
    protected static Object before;
    protected static Object after;
    protected static Object sameAsBefore;
    protected static Object emptyBefore;
    protected static Object emptyAfter;

    @BeforeClass
    public static void setUp() {
        before = new ArrayList<TestDataProvider.Person>();
        TestDataProvider.Person beforeEntry = new TestDataProvider.Person();
        beforeEntry.setName("Tom");
        beforeEntry.setAge(20);
        ((List) before).add(beforeEntry);

        TestDataProvider.Person beforeEntry2 = new TestDataProvider.Person();
        beforeEntry2.setName("Mike");
        beforeEntry2.setAge(30);
        ((List) before).add(beforeEntry2);

        TestDataProvider.Person beforeEntry3 = new TestDataProvider.Person();
        beforeEntry3.setName("Molly");
        beforeEntry3.setAge(30);
        ((List) before).add(beforeEntry3);


        after = new ArrayList<TestDataProvider.Person>();
        TestDataProvider.Person afterEntry = new TestDataProvider.Person();
        afterEntry.setName("Tom");
        afterEntry.setAge(21);

        ((List) after).add(afterEntry);


        TestDataProvider.Person afterEntry2 = new TestDataProvider.Person();
        afterEntry2.setName("Mike");
        afterEntry2.setAge(35);
        ((List) after).add(afterEntry2);

        TestDataProvider.Person afterEntry3 = new TestDataProvider.Person();
        afterEntry3.setName("Simon");
        afterEntry3.setAge(35);
        ((List) after).add(afterEntry3);

        sameAsBefore = new ArrayList<TestDataProvider.Person>();

        TestDataProvider.Person sameAsBeforeEntry = new TestDataProvider.Person();
        sameAsBeforeEntry.setName("Tom");
        sameAsBeforeEntry.setAge(20);
        ((List) sameAsBefore).add(sameAsBeforeEntry);


        emptyBefore = new ArrayList<TestDataProvider.Person>();
        emptyAfter = new ArrayList<TestDataProvider.Person>();
    }

    @Test
    public void testUpdate() {
        DiffResponse diffResponse = DiffComputeEngine.getInstance().findDifferences(before, after);
        new com.paresh.diff.renderer.ExcelRenderer().render(before, after, diffResponse);
    }

}
