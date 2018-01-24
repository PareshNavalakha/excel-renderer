import com.paresh.diff.dto.DiffResponse;
import com.paresh.diff.renderer.ExcelRenderer;
import com.paresh.diff.util.DiffComputeEngine;
import org.junit.BeforeClass;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

public class ExcelRendererTest {
    private static Object before;
    private static Object after;


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

        TestDataProvider.Person beforeEntry4 = new TestDataProvider.Person();
        beforeEntry4.setName("Same");
        beforeEntry4.setAge(20);
        ((List) before).add(beforeEntry4);

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


        TestDataProvider.Person afterEntry4 = new TestDataProvider.Person();
        afterEntry4.setName("Same");
        afterEntry4.setAge(20);
        ((List) after).add(afterEntry4);

    }

    @Test
    public void testUpdate() {
        DiffResponse diffResponse = DiffComputeEngine.getInstance().findDifferences(before, after);
        new com.paresh.diff.renderer.ExcelRenderer().render(before, after, diffResponse);
    }

}
