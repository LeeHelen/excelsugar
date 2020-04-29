package com.kyle.excel.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;

import com.kyle.excel.exception.ExcelAnalysisException;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.util.IOUtils;

/**
 * File操作工具类
 *
 * @package: com.kyle.excel
 * @className: FileUtil
 * @author: Kyle.Y.Li
 * @since 1.0.0 --4/29/2020 10:35
 * @Param 
 * @return 
 */
public class FileUtil {
    private static final int WRITE_BUFF_SIZE = 8192;

    /**
     * 获取文件名称
     *
     * @param fileFullName
     * @return 文件名称
     */
    public static String getBaseName(final String fileFullName) {
        if (StringUtils.isBlank(fileFullName)) {
            return StringUtils.EMPTY;
        }
        return FilenameUtils.getBaseName(fileFullName);
    }

    /**
     * 获取文件名称（包含后缀名）
     *
     * @param fileFullName 文件全名
     * @return 包含后缀名的文件名称
     */
    public static String getFileName(final String fileFullName) {
        if (StringUtils.isBlank(fileFullName)) {
            return StringUtils.EMPTY;
        }
        return FilenameUtils.getName(fileFullName);
    }

    /**
     * 获取文件的完整目录
     *
     * @param fileFullName 文件全名
     * @return 文件的完整目录
     */
    public static String getFullPath(final String fileFullName) {
        if (StringUtils.isBlank(fileFullName)) {
            return StringUtils.EMPTY;
        }
        return FilenameUtils.getFullPath(fileFullName);
    }

    /**
     * 获取文件后缀名
     *
     * @param fileFullName 文件全名
     * @return 文件后缀名
     */
    public static String getExtension(final String fileFullName) {
        if (StringUtils.isBlank(fileFullName)) {
            return StringUtils.EMPTY;
        }
        return FilenameUtils.getExtension(fileFullName);
    }

    /**
     * 转换路径分隔符为当前系统分隔符
     *
     * @param fileFullName 文件全名
     * @return 当前系统可识别的文件名
     */
    public static String getNormalFileFullNameInSystem(final String fileFullName) {
        if (StringUtils.isBlank(fileFullName)) {
            return StringUtils.EMPTY;
        }
        return FilenameUtils.separatorsToSystem(fileFullName);
    }

    /**
     * 判断文件扩展名是否包含在指定的扩展名集合中
     *
     * @param fileFullName 文件全名
     * @param extensions   指定的扩展名集合
     * @return
     */
    public static boolean isExtension(final String fileFullName, final String... extensions) {
        if (StringUtils.isBlank(fileFullName) || StringUtils.isAllBlank(extensions)) {
            return false;
        }
        return FilenameUtils.isExtension(fileFullName, extensions);
    }

    /**
     * 判断文件扩展名是否包含在指定的扩展名集合中（忽略大小写）
     *
     * @param fileFullName
     * @param extensions
     * @return
     */
    public static boolean isExtensionIgnoreCase(String fileFullName, String... extensions) {
        if (StringUtils.isBlank(fileFullName) || StringUtils.isAllBlank(extensions)) {
            return false;
        }
        fileFullName = fileFullName.toLowerCase();
        extensions = (String[]) Arrays.stream(extensions).map(String::toLowerCase).toArray(String[]::new);
        return isExtension(fileFullName, extensions);
    }

    /**
     * 根据文件全名创建一个File对象
     *
     * @param fileFullName
     * @return
     */
    public static File getFile(String fileFullName) {
        if (StringUtils.isBlank(fileFullName)) {
            throw new NullPointerException("fileFullName must not be null");
        }
        return new File(fileFullName);
    }

    /**
     * 根据文件名创建一个File对象
     *
     * @param fileNames
     * @return
     */
    public static File getFile(final String... fileNames) {
        if (StringUtils.isAllBlank(fileNames)) {
            throw new NullPointerException("fileNames must not be null");
        }
        return FileUtils.getFile(fileNames);
    }

    /**
     * 判断文件是否存在
     *
     * @param fileFullName
     * @return
     */
    public static boolean isExists(final String fileFullName) {
        if (StringUtils.isBlank(fileFullName)) {
            return false;
        }
        return isExists(getFile(fileFullName));
    }

    /**
     * 判断文件是否存在
     *
     * @param file
     * @return
     */
    public static boolean isExists(final File file) {
        if (file == null) {
            return false;
        }
        return file.isFile() && file.exists();
    }

    /**
     * 判断文件目录是否存在
     *
     * @param fileFullName
     * @return
     */
    public static boolean isExistsDic(final String fileFullName) {
        if (StringUtils.isBlank(fileFullName)) {
            return false;
        }
        String fileFullPath = getFullPath(fileFullName);
        return isExistsDic(getFile(fileFullPath));
    }

    /**
     * 判断文件目录是否存在
     *
     * @param file
     * @return
     */
    public static boolean isExistsDic(final File file) {
        if (file == null) {
            return false;
        }
        return file.isDirectory() && file.exists();
    }

    /**
     * 获取文件大小
     *
     * @param fileFullName
     * @return
     */
    public static double fileSizeOf(final String fileFullName) {
        if (StringUtils.isBlank(fileFullName)) {
            return 0;
        }
        return fileSizeOf(getFile(fileFullName));
    }

    /**
     * 获取文件大小
     *
     * @param file
     * @return
     */
    public static double fileSizeOf(final File file) {
        if (file == null) {
            return 0;
        }
        return FileUtils.sizeOf(file);
    }

    /**
     * 获取文件绝对路径
     *
     * @param fileFullName
     * @return
     */
    public static String getAbsolutePath(final String fileFullName) {
        if (StringUtils.isBlank(fileFullName) || !isExists(fileFullName)) {
            return fileFullName;
        }
        return getAbsolutePath(getFile(fileFullName));
    }

    /**
     * 获取文件绝对路径
     *
     * @param file
     * @return
     */
    public static String getAbsolutePath(final File file) {
        if (file == null || !isExists(file)) {
            return StringUtils.EMPTY;
        }
        return file.getAbsolutePath();
    }

    /**
     * 读取File到FileInputStream
     *
     * @param fileFullName
     * @return
     */
    public static FileInputStream openInputStream(final String fileFullName) {
        File file = getFile(fileFullName);
        return openInputStream(file);
    }

    /**
     * 读取File到FileInputStream
     *
     * @param file
     * @return
     */
    public static FileInputStream openInputStream(File file) {
        FileInputStream fileInputStream = null;
        try {
            if (file == null || !isExists(file)) {
                throw new FileNotFoundException("File '" + file + "' does not exist");
            }
            fileInputStream = FileUtils.openInputStream(file);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return fileInputStream;
    }

    /**
     * 获得文件的16进制数据
     *
     * @param buffer
     * @return
     */
    public static String getFileHexString(byte[] buffer) {
        StringBuilder stringBuilder = new StringBuilder();
        if (buffer == null || buffer.length <= 0) {
            return null;
        }
        for (int i = 0; i < buffer.length; i++) {
            int v = buffer[i] & 0xFF;
            String hv = Integer.toHexString(v);
            if (hv.length() < 2) {
                stringBuilder.append(0);
            }
            stringBuilder.append(hv);
        }
        return stringBuilder.toString();
    }

    /**
     * 转换 InputStream 为 File
     *
     * @param inputStream
     * @param file
     * @return
     */
    public static File convertInputstreamToFile(InputStream inputStream, File file) {
        try (OutputStream outputStream = new FileOutputStream(file)) {
            int bytesRead = 0;
            byte[] buffer = new byte[8192];
            while ((bytesRead = inputStream.read(buffer, 0, 8192)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            IOUtils.closeQuietly(inputStream);
        }
        return file;
    }

    /**
     * Write inputStream to file
     *
     * @param file
     * @param inputStream
     */
    public static void writeToFile(File file, InputStream inputStream) {
        OutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(file);
            int bytesRead;
            byte[] buffer = new byte[WRITE_BUFF_SIZE];
            while ((bytesRead = inputStream.read(buffer, 0, WRITE_BUFF_SIZE)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
            }
        } catch (Exception e) {
            throw new ExcelAnalysisException("Can not create temporary file!", e);
        } finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    throw new ExcelAnalysisException("Can not close 'outputStream'!", e);
                }
            }
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    throw new ExcelAnalysisException("Can not close 'inputStream'", e);
                }
            }
        }
    }

    /**
     * 转换 InputStream 为 byte[]
     *
     * @param inputStream
     * @return
     */
    public static byte[] convertInputstreamToByteArray(InputStream inputStream) {
        byte[] buffer = null;
        try {
            buffer = IOUtils.toByteArray(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            // 关闭输入流
            IOUtils.closeQuietly(inputStream);
        }
        return buffer;
    }

    /**
     * Reads the contents of a file into a byte array. * The file is always closed.
     *
     * @param file
     * @return
     * @throws IOException
     */
    public static byte[] readFileToByteArray(final File file) throws IOException {
        InputStream in = openInputStream(file);
        try {
            final long fileLength = file.length();
            return fileLength > 0 ? IOUtils.toByteArray(in, (int)fileLength) : IOUtils.toByteArray(in);
        } finally {
            in.close();
        }
    }

    /**
     * 转换 InputStream 为 outputStream
     *
     * @param inputStream
     * @param outputStream
     */
    public static void convertInputstreamToOutputStream(InputStream inputStream, OutputStream outputStream) {
        try {
            IOUtils.copy(inputStream, outputStream, 8192);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            // 关闭输入流
            IOUtils.closeQuietly(inputStream);
        }
    }
}
