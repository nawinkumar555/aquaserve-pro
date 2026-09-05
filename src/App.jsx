import { useEffect, useState } from "react";
import { auth, db } from "./firebase";
import { onAuthStateChanged } from "firebase/auth";
import { doc, getDoc } from "firebase/firestore";

// Importing the professional components
import AdminDashboard from "./components/AdminDashboard";
import EmployeeApp from "./components/EmployeeApp";
import Login from "./components/Login";

function App() {
  const [user, setUser] = useState(null);
  const [role, setRole] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      if (currentUser) {
        // --- LOGIC START: ROLE CHECKING ---
        
        // 1. Check if the user is the Admin by your specific email
        if (currentUser.email === "nawinkumarofficial@gmail.com") {
          setRole("admin");
          setUser(currentUser);
        } else {
          // 2. Otherwise, check the 'users' collection for technician details
          // This allows us to get the technician's name for attendance
          const userDoc = await getDoc(doc(db, "users", currentUser.uid));
          
          if (userDoc.exists()) {
            setRole("employee");
            setUser({ 
              uid: currentUser.uid, 
              email: currentUser.email, 
              ...userDoc.data() 
            });
          } else {
            // Fallback for staff login if no firestore doc exists yet
            setRole("employee");
            setUser(currentUser);
          }
        }
        // --- LOGIC END ---
      } else {
        setUser(null);
        setRole(null);
      }
      setLoading(false);
    });
    return () => unsubscribe();
  }, []);

  // Professional loading screen
  if (loading) {
    return (
      <div className="h-screen bg-slate-900 flex flex-col items-center justify-center">
        <div className="w-12 h-12 border-4 border-blue-500 border-t-transparent rounded-full animate-spin mb-4"></div>
        <p className="text-blue-500 font-black tracking-widest text-xs uppercase">Aquaserve OS Loading</p>
      </div>
    );
  }

  // If no user is logged in, show Login
  if (!user) {
    return <Login />;
  }

  // Show the correct dashboard based on role
  // We pass 'user' to EmployeeApp so it knows the name (Arun/Suresh)
  return role === "admin" ? <AdminDashboard /> : <EmployeeApp userDetails={user} />;
}

export default App;