package batchJob;

import java.util.Collection;
import java.util.Iterator;

import com.agile.api.*;

import connect.ConnecPLM;

public class GetClassAttributes {

	public static INode toNode(Object o) {
		return (INode) o;
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		IAgileSession myServer = ConnecPLM.logInAsAdmin();
		try {
			Object[] top_node = myServer.getAdminInstance().getNode(NodeConstants.NODE_AGILE_CLASSES).getChildren();

			for (Object Obj0 : top_node) {
				INode node0 = (INode) Obj0;
				System.out.println(node0.getName() + ", " + node0.getId());

				Object[] ObjList1 = node0.getChildren();
				for (Object Obj1 : top_node) {
					INode node1 = (INode) Obj1;
					System.out.println(node1.getName() + ", " + node1.getId());

					Object[] ObjList2 = node1.getChildren();
					for (Object o : ObjList2) {
					IProperty[] pp = (IProperty[]) toNode(o).getProperties();
					for (IProperty ppp : pp) {
						System.out.println(ppp.getName() + ", " + ppp.getValue());

					}

				}
					
					for (Object Obj2 : ObjList2) {
						INode node2 = (INode) Obj2;
						//System.out.println(node2.getName() + ", " + node2.getId());

						
//						for (Object o : ObjList2) {
//							IProperty[] pp = (IProperty[]) toNode(o).getProperties();
//							for (IProperty ppp : pp) {
//								System.out.println(ppp.getName() + ", " + ppp.getValue());
//
//							}
//
//						}

					}

				}

			}

			// ITreeNode sub_nodes = top_node.getChildNode();
//			Collection childNodes = top_node.getChildNodes();
//			for (Iterator it = childNodes.iterator(); it.hasNext();) {
//				INode node = (INode) it.next();
//				System.out.println(node.getName() + ", " + node.getId());
//}
		} catch (APIException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

	}

}
