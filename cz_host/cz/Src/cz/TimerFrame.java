package cz;

import java.awt.BorderLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.Timer;


public class TimerFrame extends JFrame implements ActionListener{
	
	
	Timer t;
	public int count = 10;
	public JLabel lb = new JLabel();
	public JLabel pl = new JLabel();
	
	
	public TimerFrame(){
		
		setTitle( "TimerFrame" );
		setSize( 300 , 100 );
		getContentPane().setLayout( new BorderLayout() );
		setDefaultCloseOperation( 3 );
		
		JButton but = new JButton( "START" );
		but.addActionListener( this );
		
		getContentPane().add( lb , BorderLayout.NORTH );
		getContentPane().add( pl , BorderLayout.CENTER );
		getContentPane().add( but , BorderLayout.SOUTH );
		
		t = new Timer( 1000 , new MyAction() );
		
		setVisible( true );
		
	}
	
	
	public void actionPerformed( ActionEvent a ){
		
		t.restart();
		
	}
	
	
	class MyAction implements ActionListener{
		
		
		int square = 0;
		String s = new String();
		
		
		public void actionPerformed( ActionEvent a ){
			
			
			count = count - 1;
			Integer i = new Integer( count );
			lb.setText( i.toString() );
			
			if ( ( count % 10 )==0 ){
				setVisible(false);
				System.exit(0);
			}
			
			pl.setText( s + "•bŒã‚É‰æ–Ê‚ð•Â‚¶‚Ü‚·");
			
		}
	}
	
	
	public static void main( String args[] ){
		
		TimerFrame tf = new TimerFrame();
		
	}
}