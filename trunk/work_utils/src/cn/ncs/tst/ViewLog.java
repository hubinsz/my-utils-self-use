package cn.ncs.tst;

import java.io.EOFException;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.net.InetSocketAddress;
import java.net.ServerSocket;
import java.net.Socket;
import java.net.SocketException;

import org.apache.log4j.PatternLayout;
import org.apache.log4j.spi.LoggingEvent;

public class ViewLog {

	private static PatternLayout pl = new PatternLayout(
			"%5p %d{yyyy/MM/dd HH:mm:ss} %m%n");

	public static String getFormattedString(String stackTrace[]) {
		StringBuffer buffer = new StringBuffer();
		if (getStackTracePresent(stackTrace)) {
			buffer.append("\nThrowable: ").append(stackTrace[0]).append('\n');
			for (int i = 1; i < stackTrace.length; i++)
				buffer.append(stackTrace[i]).append('\n');
		}
		String formattedString = buffer.toString();
		buffer.setLength(0);
		buffer = null;
		return formattedString;
	}

	public static boolean getStackTracePresent(String stackTrace[]) {
		return stackTrace != null;
	}

	public static void main(String[] args) {
		Socket socket = null;
		ObjectInputStream in = null;
		System.out.println();
		System.out.println("=======================================");
		System.out.println("* Log View Toos (author: lurenjia) *");
		System.out.println("* Ver 0.02 build 2006/12/20 *");
		System.out.println("=======================================");

		if (args.length < 2) {
			System.out.println("java LogView ");
			return;
		}
		try {
			String hostname = args[0];
			int port = Integer.parseInt(args[1]);
			InetSocketAddress isa = new InetSocketAddress(hostname, port);
			ServerSocket server = new ServerSocket();
			server.bind(isa);
			socket = server.accept();
			System.out.println("Listening " + hostname + ":" + port);
			in = new ObjectInputStream(socket.getInputStream());
			do {
				LoggingEvent loggingEvent = (LoggingEvent) in.readObject();
				System.out.print(pl.format(loggingEvent));
				String stackTrace[] = loggingEvent.getThrowableStrRep();
				System.out.print(getFormattedString(stackTrace));
			} while (true);
		} catch (EOFException e) {
		} catch (SocketException s) {
		} catch (Throwable t) {
			t.printStackTrace();
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (socket != null) {
				try {
					socket.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			in = null;
			socket = null;
		}
	}

}
