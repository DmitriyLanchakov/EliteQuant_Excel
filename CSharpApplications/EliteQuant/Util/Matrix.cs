using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EliteQuant
{
    public class MatrixFunctions
    {
        #region Matrix
        public static double[] GetMatrixRow(double[,] matrix, int idx)
        {
            int row = matrix.GetLength(0);
            int col = matrix.GetLength(1);

            double[] ret = new double[col];
            for (int i = 0; i < col; i++)
                ret[i] = matrix[idx, i];

            return ret;
        }

        public static double[] GetMatrixColumn(double[,] matrix, int idx)
        {
            int row = matrix.GetLength(0);
            int col = matrix.GetLength(1);

            double[] ret = new double[row];
            for (int i = 0; i < row; i++)
                ret[i] = matrix[i, idx];

            return ret;
        }

        public static Matrix MatrixToEQMatrix(double[,] matrix, bool transpose = false)
        {
            if (transpose)
            {
                int row = matrix.GetLength(1);
                int col = matrix.GetLength(0);

                Matrix m = new Matrix((uint)row, (uint)col);
                for (int i = 0; i < row; i++)
                {
                    for (int j = 0; j < col; j++)
                    {
                        m.set((uint)i, (uint)j, matrix[j, i]);
                    }
                }

                return m;
            }
            else
            {
                int row = matrix.GetLength(0);
                int col = matrix.GetLength(1);

                Matrix m = new Matrix((uint)row, (uint)col);
                for (int i = 0; i < row; i++)
                {
                    for (int j = 0; j < col; j++)
                    {
                        m.set((uint)i, (uint)j, matrix[i, j]);
                    }
                }

                return m;
            } 
        }

        public static EliteQuant.DoubleVector VectorToEQVector(double[] v)
        {
            int row = v.Count();

            EliteQuant.DoubleVector vec = new EliteQuant.DoubleVector(row);

            for (int i = 0; i < row; i++)
                vec[i] = v[i];

            return vec;
        }

        public static double[,] EQMatrixToMatrix(Matrix matrix, bool transpose = false)
        {
            if (transpose)
            {
                uint row = matrix.columns();
                uint col = matrix.rows();

                double[,] ret = new double[row, col];

                for (uint i = 0; i < row; i++)
                {
                    for (uint j = 0; j < col; j++)
                    {
                        ret[i, j] = matrix.get(j, i);
                    }
                }

                return ret;
            }
            else
            {
                uint row = matrix.rows();
                uint col = matrix.columns();

                double[,] ret = new double[row, col];

                for (uint i = 0; i < row; i++)
                {
                    for (uint j = 0; j < col; j++)
                    {
                        ret[i, j] = matrix.get(i, j);
                    }
                }

                return ret;
            }
        }

        public static double[] EQVectorToVector(EliteQuant.DoubleVector v)
        {
            int row = v.Count();

            double[] vec = new double[row];

            for (int i = 0; i < row; i++)
                vec[i] = v[i];

            return vec;
        }

        public static double[] EQVectorToVector(EliteQuant.QlArray v)
        {
            uint row = v.size();

            double[] vec = new double[row];

            for (uint i = 0; i < row; i++)
                vec[i] = v.get(i);

            return vec;
        }

        public static void WriteMatrixToFile(double[,] matrix, string path, string[] title=null)
        {
            List<string> output = new List<string>();

            if (title != null)
                output.Add(string.Join(",", title));

            for (int i = 0; i < matrix.GetLength(0); i++)
            {
                output.Add(string.Join(",", GetMatrixRow(matrix, i)));
            }

            System.IO.File.WriteAllLines(path, output);
        }

        public static void WriteMatrixToFile(List<List<double>> matrix, string path, string[] title = null)
        {
            List<string> output = new List<string>();

            if (title != null)
                output.Add(string.Join(",", title));

            for (int i = 0; i < matrix.Count; i++)
            {
                output.Add(string.Join(",", matrix[i]));
            }

            System.IO.File.WriteAllLines(path, output);
        }

        public static void WriteVectorToFile(double[] vector, string path)
        {
            List<string> output = new List<string>();

            for (int i = 0; i < vector.Count(); i++)
            {
                output.Add(vector[i].ToString());
            }

            System.IO.File.WriteAllLines(path, output);
        }

        public static void WriteVectorToFile(List<string> vector, string path)
        {
            List<string> output = new List<string>();

            for (int i = 0; i < vector.Count(); i++)
            {
                output.Add(vector[i].ToString());
            }

            System.IO.File.WriteAllLines(path, output);
        }

        public static double[,] ReadMatrixFromFile(string path)
        {
            string[] input = System.IO.File.ReadAllLines(path);
            string[] row = input[0].Split(',');
            double[,] ret = new double[input.Length, row.Length];
            
            try
            {
                for (int i = 0; i < ret.GetLength(0); i++)
                {
                    for (int j = 0; j < ret.GetLength(1); j++)
                    {
                        row = input[i].Split(',');
                        ret[i, j] = Convert.ToDouble(row[j]);
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }

            return ret;
        }

        public static double[] ReadVectorFromFile(string path)
        {
            return null;
        }

        public static double[,] MultiplyMatrix(double[,] A, double[,] B)
        {
            int rA = A.GetLength(0);
            int cA = A.GetLength(1);
            int rB = B.GetLength(0);
            int cB = B.GetLength(1);
            double temp = 0;
            double[,] kHasil = new double[rA, cB];
            if (cA != rB)
            {
                throw new ArgumentException("matrices can't be multiplied !!");
            }
            else
            {
                for (int i = 0; i < rA; i++)
                {
                    for (int j = 0; j < cB; j++)
                    {
                        temp = 0;
                        for (int k = 0; k < cA; k++)
                        {
                            temp += A[i, k] * B[k, j];
                        }
                        kHasil[i, j] = temp;
                    }
                }
                return kHasil;
            }
        }

        public static double[,] AddMatrix(double[,] A, double[,] B)
        {
            int rA = A.GetLength(0);
            int cA = A.GetLength(1);
            int rB = B.GetLength(0);
            int cB = B.GetLength(1);
            
            if ((rA != rB) && (cA != cB))
            {
                throw new ArgumentException("matrices can't be added !!");
            }
            else
            {
                double[,] kHasil = new double[rA, cA];

                for (int i = 0; i < rA; i++)
                {
                    for (int j = 0; j < cA; j++)
                    {
                        kHasil[i, j] = A[i, j] + B[i, j];
                    }
                }
                return kHasil;
            }
        }

        public static double[,] TransposeMatrix(double[,] A)
        {
            int rA = A.GetLength(0);
            int cA = A.GetLength(1);
            
            double[,] ret = new double[cA, rA];
            
            for (int i = 0; i < rA; i++)
            {
                for (int j = 0; j < cA; j++)
                {
                    ret[j, i] = A[i, j];
                }
            }

            return ret;
            
        }
        #endregion
    }
}
