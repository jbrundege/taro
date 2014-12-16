package taro;

public class TaroException extends RuntimeException {

    public TaroException() {
    }

    public TaroException(String message) {
        super(message);
    }

    public TaroException(String message, Throwable cause) {
        super(message, cause);
    }

    public TaroException(Throwable cause) {
        super(cause);
    }

    public TaroException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }

}
